#!/usr/bin/env python3
"""
Droplet Counter GUI V6.1

Purpose
-------
A GUI for counting droplets in microscopy image series, optimized for dark
brightfield-like droplets on a brighter background.

Main changes in V6.1
--------------------
- Preview performance optimization: cached signal images and cached detection results.
- Marker show/hide, label display, rejected-candidate display, and preview gamma no longer re-run detection.
- Expensive shape metrics are now calculated only after cheap intensity/size gates pass.
- Slider updates use a longer debounce to avoid re-analyzing continuously while dragging.

Main changes in V6
------------------
- Main tab reduced to three linked routine controls: sensitivity, shape strictness, and close-droplet separation.
- Added shape-first acceptance filters: roundness, circularity, eccentricity, angular edge coverage, and angular uniformity.
- Added close-droplet splitting using local peak detection before duplicate removal.
- Advanced tab now exposes the detailed shape, splitting, blank-guard, size, and signal parameters.
- Excel export records shape-filter diagnostics and splitting diagnostics.

Main changes in V5.3
--------------------
- Blank-frame guard uses droplet roundness as its main protection criterion.
- Frames are not blank-guarded when they contain enough convincing round droplets.
- Excel export records roundness metrics and blank-guard diagnostics.

Main changes in V5.2
--------------------
- Excel export includes a temperature_C column extracted from filenames such as 25_76C_2W-oT_cropped.
- temperature_C is placed after frame_name in per_droplet and per_frame_summary.

Main changes in V5.1
--------------------
- Added a preview toolbar button that hides or shows marker overlays without changing detection or export.
- Marker visibility is preview-only: counts, measurements, presets, and Excel output are unchanged.
- Press M as a shortcut to toggle preview markers.

Main changes in V4.1
--------------------
- Main and Advanced controls are explicitly linked by default.
- Droplet sensitivity and false-positive rejection now move visible effective Advanced thresholds.
- Linked mode can be disabled for manual expert tuning.

Main changes in V4
------------------
- All right-side tabs are scrollable.
- Main detection panel has only two key parameters:
    1. Droplet sensitivity
    2. False-positive rejection
- Advanced parameters are moved to a separate Advanced tab.
- Parameter sets can be saved, loaded, deleted, and reused.
- Parameters can be locked to avoid accidental changes.
- Excel export records all parameters, selected preset, and software version.
- Default detector is tuned for the representative images supplied by the user:
    dark in-focus droplets are detected,
    weak debris-only frames are suppressed by a blank-frame guard.

Recommended install
-------------------
Run this once in Command Prompt or PowerShell:

    pip install numpy pillow tifffile pandas openpyxl scikit-image scipy

Optional OCR support:

    pip install pytesseract

OCR also requires the external Tesseract OCR application.
Manual scale calibration does not require OCR.
"""

from __future__ import annotations

import json
import math
import re
import sys
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import numpy as np

try:
    from PIL import Image, ImageDraw, ImageTk
except Exception as exc:
    raise RuntimeError("Pillow is required. Install it with: pip install pillow") from exc

try:
    import pandas as pd
except Exception as exc:
    raise RuntimeError(
        "pandas and openpyxl are required for Excel export. Install with: pip install pandas openpyxl"
    ) from exc

try:
    import tifffile
except Exception:
    tifffile = None

try:
    from scipy import ndimage as ndi
except Exception as exc:
    raise RuntimeError("scipy is required. Install it with: pip install scipy") from exc

try:
    from skimage import feature, filters, measure, morphology, segmentation
except Exception as exc:
    raise RuntimeError("scikit-image is required. Install it with: pip install scikit-image") from exc


SOFTWARE_VERSION = "Droplet Counter GUI V6.1 - fast cached shape-first detector"
SUPPORTED_EXTENSIONS = {".tif", ".tiff", ".png", ".jpg", ".jpeg", ".bmp"}


@dataclass
class FrameRecord:
    source_path: Path
    page_index: int
    name: str
    image: np.ndarray


@dataclass
class DetectionOutput:
    overlay_image: Optional[Image.Image]
    rows: List[Dict[str, object]]
    rejected_rows: List[Dict[str, object]]
    count: int
    frame_metrics: Dict[str, object]


def natural_sort_key(path: Path):
    text = str(path.name)
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", text)]


def extract_temperature_c_from_name(name: str) -> Optional[float]:
    """
    Extract temperature in Celsius from filenames/frame labels.

    Expected examples:
    - 25_76C_2W-oT_cropped -> 76
    - 25_76.5C_2W-oT_cropped -> 76.5
    - frame_80°C_sample -> 80

    Returns None when no temperature token is found.
    """
    text = str(name)
    # Remove common image extensions and TIFF-stack page labels only for cleaner matching.
    text = re.sub(r"\.(tif|tiff|png|jpg|jpeg|bmp)(?:\s*\[\d+\])?$", "", text, flags=re.IGNORECASE)

    # Match a standalone number followed immediately or near-immediately by C/°C.
    # The negative lookbehind prevents treating 25_76C as 25.76C; it starts at 76C after the underscore.
    pattern = r"(?<![A-Za-z0-9])([-+]?\d+(?:[\.,]\d+)?)\s*°?\s*C(?![A-Za-z])"
    matches = list(re.finditer(pattern, text, flags=re.IGNORECASE))
    if not matches:
        return None

    raw = matches[-1].group(1).replace(",", ".")
    try:
        return float(raw)
    except Exception:
        return None


def to_grayscale_array(image: np.ndarray) -> np.ndarray:
    arr = np.asarray(image)

    if arr.ndim == 3 and arr.shape[-1] in (3, 4):
        rgb = arr[..., :3].astype(np.float32)
        gray = 0.2126 * rgb[..., 0] + 0.7152 * rgb[..., 1] + 0.0722 * rgb[..., 2]
    elif arr.ndim > 2:
        gray = np.max(arr, axis=0).astype(np.float32)
    else:
        gray = arr.astype(np.float32)

    gray = np.nan_to_num(gray, nan=0.0, posinf=0.0, neginf=0.0)
    return gray


def robust_normalize(image: np.ndarray, low_percentile: float = 0.2, high_percentile: float = 99.8) -> np.ndarray:
    arr = to_grayscale_array(image)
    lo, hi = np.percentile(arr, [low_percentile, high_percentile])

    if hi <= lo:
        lo = float(np.min(arr))
        hi = float(np.max(arr))

    if hi <= lo:
        return np.zeros(arr.shape, dtype=np.float32)

    out = (arr - lo) / (hi - lo)
    return np.clip(out, 0.0, 1.0).astype(np.float32)


def display_to_float01(image: np.ndarray, gamma: float = 0.75) -> np.ndarray:
    out = robust_normalize(image, 0.1, 99.9)

    try:
        gamma = float(gamma)
    except Exception:
        gamma = 0.75

    gamma = min(max(gamma, 0.20), 3.00)
    out = out ** gamma
    return out.astype(np.float32)


def load_image_file(path: Path) -> List[FrameRecord]:
    frames: List[FrameRecord] = []
    suffix = path.suffix.lower()

    if suffix in {".tif", ".tiff"} and tifffile is not None:
        try:
            arr = tifffile.imread(str(path))
        except Exception:
            arr = None

        if arr is not None:
            arr = np.asarray(arr)

            if arr.ndim == 2:
                frames.append(FrameRecord(path, 0, path.name, arr))
                return frames

            if arr.ndim == 3 and arr.shape[-1] in (3, 4):
                frames.append(FrameRecord(path, 0, path.name, arr))
                return frames

            if arr.ndim == 3:
                for i in range(arr.shape[0]):
                    frames.append(FrameRecord(path, i, f"{path.name} [{i + 1}]", arr[i]))
                return frames

            if arr.ndim == 4 and arr.shape[-1] in (3, 4):
                for i in range(arr.shape[0]):
                    frames.append(FrameRecord(path, i, f"{path.name} [{i + 1}]", arr[i]))
                return frames

            projected = np.max(arr, axis=0)
            frames.append(FrameRecord(path, 0, path.name, projected))
            return frames

    with Image.open(path) as img:
        page = 0
        while True:
            try:
                img.seek(page)
                arr = np.array(img)
                label = f"{path.name} [{page + 1}]" if page else path.name
                frames.append(FrameRecord(path, page, label, arr))
                page += 1
            except EOFError:
                break

    return frames


def detect_scale_bar_pixels(image: np.ndarray) -> Optional[float]:
    gray = robust_normalize(image)
    h, w = gray.shape[:2]
    y0 = int(h * 0.55)
    crop = gray[y0:h, :]

    candidates = []

    def find_candidates(binary: np.ndarray, mode_name: str):
        cleaned = morphology.remove_small_objects(binary.astype(bool), min_size=max(10, int(w * 0.002)))
        lab = measure.label(cleaned)

        for prop in measure.regionprops(lab):
            minr, minc, maxr, maxc = prop.bbox
            height = maxr - minr
            width = maxc - minc

            if width < max(15, int(w * 0.02)):
                continue
            if height < 1 or height > max(20, int(h * 0.04)):
                continue
            if width / max(height, 1) < 5:
                continue
            if width > w * 0.8:
                continue

            y_center = y0 + (minr + maxr) / 2
            if y_center < h * 0.6:
                continue

            score = width * 2.0 + width / max(height, 1) + prop.area * 0.02 + y_center * 0.01
            candidates.append((score, float(width), mode_name))

    bright_thr = np.percentile(crop, 99.2)
    dark_thr = np.percentile(crop, 0.8)

    find_candidates(crop >= bright_thr, "bright")
    find_candidates(crop <= dark_thr, "dark")

    if not candidates:
        return None

    candidates.sort(reverse=True, key=lambda x: x[0])
    return candidates[0][1]


def ocr_scale_text(image: np.ndarray) -> Optional[float]:
    try:
        import pytesseract
    except Exception:
        return None

    gray = robust_normalize(image)
    h, _w = gray.shape[:2]
    crop = gray[int(h * 0.55):h, :]
    crop8 = (np.clip(crop, 0, 1) * 255).astype(np.uint8)

    pil = Image.fromarray(crop8)
    pil = pil.resize((pil.width * 3, pil.height * 3), Image.Resampling.BICUBIC)

    texts = []
    for candidate in [pil, Image.fromarray(255 - np.array(pil))]:
        try:
            texts.append(pytesseract.image_to_string(candidate, config="--psm 6"))
        except Exception:
            pass

    joined = "\n".join(texts)
    patterns = [
        r"(\d+(?:\.\d+)?)\s*(?:u|µ|μ)\s*m",
        r"(\d+(?:\.\d+)?)\s*microm",
        r"(\d+(?:\.\d+)?)\s*micro",
    ]

    for pat in patterns:
        m = re.search(pat, joined, flags=re.IGNORECASE)
        if m:
            try:
                val = float(m.group(1))
                if 0 < val < 100000:
                    return val
            except Exception:
                pass

    nums = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", joined)]
    plausible = [x for x in nums if 0 < x < 10000]
    if len(plausible) == 1:
        return plausible[0]

    return None


class ScrollableFrame(ttk.Frame):
    """
    Scrollable container used in every right-side tab.
    This prevents controls from being cut off on smaller screens.
    """
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)

        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        self.v_scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)

        self.inner_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.v_scrollbar.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.v_scrollbar.grid(row=0, column=1, sticky="ns")

        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)
        self.inner.bind("<Enter>", self._bind_mousewheel)
        self.inner.bind("<Leave>", self._unbind_mousewheel)

    def _on_inner_configure(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event=None):
        self.canvas.itemconfigure(self.inner_id, width=event.width)

    def _bind_mousewheel(self, event=None):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel_linux)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel_linux)

    def _unbind_mousewheel(self, event=None):
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_mousewheel_linux(self, event):
        if event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")


class DropletCounterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Droplet Counter GUI V6")
        self.root.geometry("1320x820")
        self.root.minsize(950, 620)

        self.frames: List[FrameRecord] = []
        self.current_index = 0
        self.preview_photo = None
        self.current_detection: Optional[DetectionOutput] = None
        self.current_detection_cache_key = None
        self.update_after_id: Optional[str] = None

        # Performance caches. V6 shape-first analysis is relatively expensive, so
        # preview rendering is separated from detection. Display-only changes reuse
        # current_detection instead of re-running the detector.
        self.signal_cache: Dict[tuple, Tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray]] = {}
        self.analysis_cache: Dict[tuple, Tuple[List[Dict[str, object]], List[Dict[str, object]], Dict[str, object]]] = {}
        self.max_signal_cache_entries = 20
        self.max_analysis_cache_entries = 80

        self.parameter_widgets: List[tk.Widget] = []
        self.linked_advanced_widgets: List[tk.Widget] = []
        self.current_preset_name: Optional[str] = None
        self._syncing_params = False

        self.scale_um_var = tk.StringVar(value="")
        self.scale_px_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Load a folder or a single file to begin.")
        self.frame_label_var = tk.StringVar(value="No image loaded")
        self.pixel_size_label_var = tk.StringVar(value="Pixel size: not calibrated")

        self.preset_name_var = tk.StringVar(value="")
        self.preset_select_var = tk.StringVar(value="")
        self.parameters_locked_var = tk.BooleanVar(value=False)
        self.effective_thresholds_var = tk.StringVar(value="")

        # Preview-only display toggle. This does not change detection, measurements, presets, or export.
        self.preview_markers_visible_var = tk.BooleanVar(value=True)
        self.preview_marker_button_text_var = tk.StringVar(value="Hide markers")

        self.params = {
            # Main controls. These three routine controls drive the linked Advanced settings.
            "droplet_sensitivity": tk.DoubleVar(value=70.0),
            "droplet_shape_strictness": tk.DoubleVar(value=70.0),
            "close_droplet_separation": tk.DoubleVar(value=65.0),
            # Kept for compatibility with V4/V5 presets. In V6 it follows droplet_shape_strictness.
            "false_positive_rejection": tk.DoubleVar(value=70.0),
            "link_main_to_advanced": tk.BooleanVar(value=True),

            # Effective advanced thresholds controlled by the Main sliders when linked.
            "candidate_z_threshold": tk.DoubleVar(value=1.76),
            "quality_threshold": tk.DoubleVar(value=2.44),
            "min_contrast_required": tk.DoubleVar(value=0.033),
            "min_edge_required": tk.DoubleVar(value=0.027),

            # Shape-first acceptance filters.
            "use_shape_filters": tk.BooleanVar(value=True),
            "min_roundness_required": tk.DoubleVar(value=0.64),
            "min_shape_circularity_required": tk.DoubleVar(value=0.48),
            "max_shape_eccentricity_allowed": tk.DoubleVar(value=0.64),
            "min_angular_coverage_required": tk.DoubleVar(value=0.56),
            "min_angular_uniformity_required": tk.DoubleVar(value=0.18),

            # Close-droplet splitting.
            "split_close_droplets": tk.BooleanVar(value=True),
            "split_min_diameter_px": tk.DoubleVar(value=8.0),
            "split_peak_threshold": tk.DoubleVar(value=1.95),
            "split_peak_distance_fraction": tk.DoubleVar(value=0.46),
            "split_child_radius_fraction": tk.DoubleVar(value=0.50),
            "split_max_children": tk.DoubleVar(value=6.0),

            # Core detection settings.
            "bright_objects": tk.BooleanVar(value=False),
            "preview_gamma": tk.DoubleVar(value=0.75),

            # Advanced detector settings.
            "min_diameter_px": tk.DoubleVar(value=4.0),
            "max_diameter_px": tk.DoubleVar(value=20.0),
            "background_sigma": tk.DoubleVar(value=18.0),
            "signal_blur_sigma": tk.DoubleVar(value=0.45),
            "max_overlap_fraction": tk.DoubleVar(value=0.65),
            "nms_overlap_factor": tk.DoubleVar(value=0.62),
            "blank_frame_guard": tk.BooleanVar(value=True),
            "blank_frame_score_threshold": tk.DoubleVar(value=4.0),
            "blank_guard_roundness_threshold": tk.DoubleVar(value=0.58),
            "blank_guard_min_round_droplets": tk.DoubleVar(value=1.0),
            "blank_guard_use_roundness": tk.BooleanVar(value=True),
            "exclude_border": tk.BooleanVar(value=True),
            "show_labels": tk.BooleanVar(value=False),
            "show_rejected_candidates": tk.BooleanVar(value=False),
        }

        self._build_gui()
        self._bind_param_updates()
        self.sync_advanced_from_main(refresh=False)
        self.refresh_preset_list()
        self.update_lock_state()

    @property
    def preset_file(self) -> Path:
        folder = Path.home() / "DropletCounterGUI"
        folder.mkdir(parents=True, exist_ok=True)
        return folder / "presets_v4.json"

    def _build_gui(self):
        main = ttk.Frame(self.root, padding=8)
        main.pack(fill=tk.BOTH, expand=True)

        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=0, minsize=430)
        main.rowconfigure(0, weight=1)

        left = ttk.Frame(main)
        left.grid(row=0, column=0, sticky="nsew")

        right = ttk.Frame(main, width=430)
        right.grid(row=0, column=1, sticky="ns", padx=(8, 0))
        right.grid_propagate(False)

        self._build_preview_panel(left)
        self._build_control_panel(right)
        self._build_status_bar()

    def _build_preview_panel(self, parent):
        parent.rowconfigure(1, weight=1)
        parent.columnconfigure(0, weight=1)

        top = ttk.Frame(parent)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 6))

        ttk.Button(top, text="Select Folder", command=self.load_folder).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(top, text="Select Single File", command=self.load_single_file).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(top, text="Export Excel", command=self.export_excel).pack(side=tk.LEFT, padx=(0, 6))
        self.preview_marker_button = ttk.Button(
            top,
            textvariable=self.preview_marker_button_text_var,
            command=self.toggle_preview_markers,
        )
        self.preview_marker_button.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(top, textvariable=self.frame_label_var).pack(side=tk.LEFT, padx=(12, 0))

        canvas_frame = ttk.Frame(parent, relief=tk.SUNKEN)
        canvas_frame.grid(row=1, column=0, sticky="nsew")

        self.canvas = tk.Canvas(canvas_frame, bg="black", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind("<Configure>", lambda event: self.refresh_preview(debounce=True))

        nav = ttk.Frame(parent)
        nav.grid(row=2, column=0, sticky="ew", pady=(6, 0))

        ttk.Button(nav, text="Previous", command=self.prev_frame).pack(side=tk.LEFT)
        ttk.Button(nav, text="Next", command=self.next_frame).pack(side=tk.LEFT, padx=(6, 12))

        self.frame_slider = ttk.Scale(nav, from_=0, to=0, orient=tk.HORIZONTAL, command=self.on_slider_change)
        self.frame_slider.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.root.bind("<Left>", lambda event: self.prev_frame())
        self.root.bind("<Right>", lambda event: self.next_frame())
        self.root.bind("<m>", lambda event: self.toggle_preview_markers())
        self.root.bind("<M>", lambda event: self.toggle_preview_markers())

    def toggle_preview_markers(self):
        visible = not bool(self.preview_markers_visible_var.get())
        self.preview_markers_visible_var.set(visible)
        self.preview_marker_button_text_var.set("Hide markers" if visible else "Show markers")
        self.render_current_detection_to_canvas()
        self.status_var.set("Preview markers shown." if visible else "Preview markers hidden. Detection and export are unchanged.")

    def _build_control_panel(self, parent):
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)

        notebook = ttk.Notebook(parent)
        notebook.grid(row=0, column=0, sticky="nsew")

        tabs = {}
        for name in ["Main", "Advanced", "Scale", "Presets", "Export", "Help"]:
            scroll = ScrollableFrame(notebook)
            scroll.inner.configure(padding=8)
            notebook.add(scroll, text=name)
            tabs[name] = scroll.inner

        self._build_main_tab(tabs["Main"])
        self._build_advanced_tab(tabs["Advanced"])
        self._build_scale_tab(tabs["Scale"])
        self._build_presets_tab(tabs["Presets"])
        self._build_export_tab(tabs["Export"])
        self._build_help_tab(tabs["Help"])

    def _slider_row(self, parent, label, var, from_, to, explanation, is_param=True):
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=(8, 0))

        header = ttk.Frame(row)
        header.pack(fill=tk.X)

        ttk.Label(header, text=label).pack(side=tk.LEFT)
        value_label = ttk.Label(header, textvariable=var, width=8)
        value_label.pack(side=tk.RIGHT)

        scale = ttk.Scale(row, from_=from_, to=to, variable=var, orient=tk.HORIZONTAL)
        scale.pack(fill=tk.X)

        ttk.Label(row, text=explanation, wraplength=380, foreground="#555555").pack(fill=tk.X)

        if is_param:
            self.parameter_widgets.append(scale)

        return scale

    def _entry_row(self, parent, label, var, explanation="", is_param=True):
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(row, text=label).pack(anchor=tk.W)
        entry = ttk.Entry(row, textvariable=var)
        entry.pack(fill=tk.X)
        if explanation:
            ttk.Label(row, text=explanation, wraplength=380, foreground="#555555").pack(fill=tk.X)

        if is_param:
            self.parameter_widgets.append(entry)

        return entry

    def _checkbutton_row(self, parent, label, var, explanation="", is_param=True):
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=(8, 0))
        cb = ttk.Checkbutton(row, text=label, variable=var)
        cb.pack(anchor=tk.W)
        if explanation:
            ttk.Label(row, text=explanation, wraplength=380, foreground="#555555").pack(fill=tk.X)

        if is_param:
            self.parameter_widgets.append(cb)

        return cb

    def _build_main_tab(self, parent):
        ttk.Label(
            parent,
            text=(
                "Routine controls. These three sliders drive the detailed Advanced settings: "
                "candidate detection, shape rejection, close-droplet splitting, duplicate removal, and blank-frame guard."
            ),
            wraplength=390,
        ).pack(fill=tk.X)

        self._slider_row(
            parent,
            "Droplet sensitivity",
            self.params["droplet_sensitivity"],
            1,
            100,
            "Higher values find weaker/smaller droplets. Lower values are safer when background debris is being detected.",
        )

        self._slider_row(
            parent,
            "Shape strictness",
            self.params["droplet_shape_strictness"],
            1,
            100,
            "Higher values require cleaner, rounder, sharper droplets. This is the main false-positive protector.",
        )

        self._slider_row(
            parent,
            "Close-droplet separation",
            self.params["close_droplet_separation"],
            1,
            100,
            "Higher values try harder to split neighboring droplets instead of counting them as one large blob.",
        )

        ttk.Label(parent, textvariable=self.effective_thresholds_var, wraplength=390, foreground="#333333").pack(fill=tk.X, pady=(8, 0))

        ttk.Separator(parent).pack(fill=tk.X, pady=12)

        self._checkbutton_row(
            parent,
            "Detect bright droplets instead of dark droplets",
            self.params["bright_objects"],
            "Leave unchecked for your usual dark droplets. Check only for bright fluorescence-like droplets.",
        )

        self._slider_row(
            parent,
            "Preview brightness gamma",
            self.params["preview_gamma"],
            0.2,
            3.0,
            "Only changes display brightness. It does not affect detection or export.",
        )

        self._checkbutton_row(
            parent,
            "Show droplet numbers on preview",
            self.params["show_labels"],
            "Shows the droplet_number used in the Excel export.",
        )

        lock_box = ttk.LabelFrame(parent, text="Parameter lock", padding=8)
        lock_box.pack(fill=tk.X, pady=(14, 0))
        ttk.Checkbutton(
            lock_box,
            text="Lock current parameters",
            variable=self.parameters_locked_var,
            command=self.update_lock_state,
        ).pack(anchor=tk.W)
        ttk.Label(
            lock_box,
            text="Locking prevents accidental parameter changes while you scroll through image series or export.",
            wraplength=360,
            foreground="#555555",
        ).pack(fill=tk.X)

        ttk.Label(
            parent,
            text=(
                "Suggested workflow: tune these three sliders on representative frames, save the parameter set in "
                "the Presets tab, lock it, then apply it to all image series. Use Advanced only for edge cases."
            ),
            wraplength=390,
            foreground="#333333",
        ).pack(fill=tk.X, pady=(14, 0))

    def _build_advanced_tab(self, parent):
        ttk.Label(
            parent,
            text=(
                "Advanced controls. In linked mode, the three Main sliders control these detailed parameters. "
                "Uncheck linked mode for manual expert tuning."
            ),
            wraplength=390,
        ).pack(fill=tk.X)

        self._checkbutton_row(
            parent,
            "Link Main sliders to Advanced settings",
            self.params["link_main_to_advanced"],
            "Recommended on. Off means these detailed thresholds can be tuned manually.",
        )

        ttk.Label(parent, textvariable=self.effective_thresholds_var, wraplength=390, foreground="#333333").pack(fill=tk.X, pady=(8, 0))

        ttk.Separator(parent).pack(fill=tk.X, pady=12)
        ttk.Label(parent, text="1. Candidate and intensity gates", font=("TkDefaultFont", 9, "bold")).pack(anchor=tk.W)

        for widget in [
            self._slider_row(
                parent,
                "Candidate detection threshold",
                self.params["candidate_z_threshold"],
                0.75,
                4.0,
                "Lower values are more sensitive and create more raw droplet candidates.",
            ),
            self._slider_row(
                parent,
                "Acceptance score threshold",
                self.params["quality_threshold"],
                0.8,
                4.2,
                "Higher values reject weak or ambiguous candidates after candidate detection.",
            ),
            self._slider_row(
                parent,
                "Minimum local contrast",
                self.params["min_contrast_required"],
                0,
                0.12,
                "Minimum brightness difference between droplet interior and local background.",
            ),
            self._slider_row(
                parent,
                "Minimum edge sharpness",
                self.params["min_edge_required"],
                0,
                0.12,
                "Minimum boundary strength. Higher values reject fuzzy/out-of-focus objects.",
            ),
        ]:
            self.linked_advanced_widgets.append(widget)

        ttk.Separator(parent).pack(fill=tk.X, pady=12)
        ttk.Label(parent, text="2. Shape-first droplet filters", font=("TkDefaultFont", 9, "bold")).pack(anchor=tk.W)

        self._checkbutton_row(
            parent,
            "Use shape-first filters",
            self.params["use_shape_filters"],
            "Recommended on. A candidate must be round, compact, and have edge signal around most of the circumference.",
        )

        for widget in [
            self._slider_row(
                parent,
                "Minimum roundness score",
                self.params["min_roundness_required"],
                0.20,
                0.95,
                "Combined score from circularity, eccentricity, solidity, size match, and angular edge coverage.",
            ),
            self._slider_row(
                parent,
                "Minimum object circularity",
                self.params["min_shape_circularity_required"],
                0.05,
                0.95,
                "Rejects elongated or irregular masks. Higher values require a more circle-like object outline.",
            ),
            self._slider_row(
                parent,
                "Maximum object eccentricity",
                self.params["max_shape_eccentricity_allowed"],
                0.20,
                0.98,
                "Rejects elongated objects. Lower values are stricter and closer to a perfect circle.",
            ),
            self._slider_row(
                parent,
                "Minimum angular edge coverage",
                self.params["min_angular_coverage_required"],
                0.05,
                1.00,
                "Requires droplet-like contrast around most angular sectors, not just one bright/dark side.",
            ),
            self._slider_row(
                parent,
                "Minimum angular uniformity",
                self.params["min_angular_uniformity_required"],
                0.00,
                0.90,
                "Rejects candidates with very uneven boundary signal around the circumference.",
            ),
        ]:
            self.linked_advanced_widgets.append(widget)

        ttk.Separator(parent).pack(fill=tk.X, pady=12)
        ttk.Label(parent, text="3. Close-droplet splitting", font=("TkDefaultFont", 9, "bold")).pack(anchor=tk.W)

        self._checkbutton_row(
            parent,
            "Split close droplets before duplicate removal",
            self.params["split_close_droplets"],
            "Recommended on. Large blob candidates are inspected for multiple local centers and split when possible.",
        )

        for widget in [
            self._slider_row(
                parent,
                "Minimum diameter to attempt splitting, pixels",
                self.params["split_min_diameter_px"],
                4,
                40,
                "Only candidates larger than this are tested for multiple droplet centers.",
            ),
            self._slider_row(
                parent,
                "Split peak threshold",
                self.params["split_peak_threshold"],
                0.50,
                5.00,
                "Lower values split more aggressively; higher values require stronger separate centers.",
            ),
            self._slider_row(
                parent,
                "Split peak distance fraction",
                self.params["split_peak_distance_fraction"],
                0.20,
                1.00,
                "Lower values allow closer neighboring peaks to be split.",
            ),
            self._slider_row(
                parent,
                "Split child radius fraction",
                self.params["split_child_radius_fraction"],
                0.25,
                0.80,
                "Estimated radius of each split child as a fraction of the original candidate radius.",
            ),
            self._slider_row(
                parent,
                "Maximum split children",
                self.params["split_max_children"],
                2,
                12,
                "Safety cap for how many droplets one large candidate can split into.",
            ),
        ]:
            self.linked_advanced_widgets.append(widget)

        ttk.Separator(parent).pack(fill=tk.X, pady=12)
        ttk.Label(parent, text="4. Size, background and duplicate removal", font=("TkDefaultFont", 9, "bold")).pack(anchor=tk.W)

        for widget in [
            self._slider_row(
                parent,
                "Minimum droplet diameter, pixels",
                self.params["min_diameter_px"],
                2,
                30,
                "Rejects small dots. Increase this to avoid counting tiny aggregates.",
            ),
            self._slider_row(
                parent,
                "Maximum droplet diameter, pixels",
                self.params["max_diameter_px"],
                4,
                80,
                "Rejects large aggregates and merged/non-droplet structures.",
            ),
            self._slider_row(
                parent,
                "Background sigma",
                self.params["background_sigma"],
                4,
                80,
                "Scale used to estimate the local background. Larger values are better for larger droplets.",
            ),
            self._slider_row(
                parent,
                "Signal blur sigma",
                self.params["signal_blur_sigma"],
                0,
                3,
                "Smoothing before detection. Lower values preserve close droplets; higher values suppress noise.",
            ),
            self._slider_row(
                parent,
                "Blob overlap fraction",
                self.params["max_overlap_fraction"],
                0.1,
                0.9,
                "Higher values allow more overlapping raw candidates, useful for nearby droplets.",
            ),
            self._slider_row(
                parent,
                "Duplicate-removal distance",
                self.params["nms_overlap_factor"],
                0.2,
                1.5,
                "Lower values keep nearby detections separate. Higher values remove nearby duplicates more aggressively.",
            ),
        ]:
            self.linked_advanced_widgets.append(widget)

        ttk.Separator(parent).pack(fill=tk.X, pady=12)
        ttk.Label(parent, text="5. Blank-frame guard", font=("TkDefaultFont", 9, "bold")).pack(anchor=tk.W)

        self._checkbutton_row(
            parent,
            "Blank-frame guard",
            self.params["blank_frame_guard"],
            "Suppresses all detections when the frame looks like debris/background rather than droplets.",
        )

        for widget in [
            self._slider_row(
                parent,
                "Blank-frame score threshold",
                self.params["blank_frame_score_threshold"],
                2.0,
                7.0,
                "Higher values make the blank-frame guard more aggressive.",
            ),
            self._slider_row(
                parent,
                "Blank-guard roundness threshold",
                self.params["blank_guard_roundness_threshold"],
                0.30,
                0.95,
                "Round candidates above this value protect a frame from blank guarding.",
            ),
            self._slider_row(
                parent,
                "Round droplets needed to protect frame",
                self.params["blank_guard_min_round_droplets"],
                1,
                10,
                "Usually keep at 1. Increase only if a single round artifact appears in blank frames.",
            ),
        ]:
            self.linked_advanced_widgets.append(widget)

        self._checkbutton_row(
            parent,
            "Use roundness protection for blank guard",
            self.params["blank_guard_use_roundness"],
            "Recommended on. Convincing round droplets protect a frame; non-round debris does not.",
        )

        ttk.Separator(parent).pack(fill=tk.X, pady=12)

        self._checkbutton_row(
            parent,
            "Exclude border-touching detections",
            self.params["exclude_border"],
            "Rejects partial droplets at the image edge.",
        )

        self._checkbutton_row(
            parent,
            "Show rejected candidates",
            self.params["show_rejected_candidates"],
            "For diagnostics only. Rejected candidates are shown in orange.",
        )

        ttk.Label(
            parent,
            text=(
                "V6 is intended to be shape-first: a droplet should be a clean, compact, round object with "
                "edge signal around most of its circumference. The old score-only logic is still recorded, but no longer the only gate."
            ),
            wraplength=390,
            foreground="#555555",
        ).pack(fill=tk.X, pady=(12, 0))

    def _build_scale_tab(self, parent):
        ttk.Label(
            parent,
            text=(
                "Scale calibration converts pixel measurements into micrometers. "
                "Manual calibration is more reliable than OCR."
            ),
            wraplength=390,
        ).pack(fill=tk.X)

        self._entry_row(
            parent,
            "Scale-bar length, micrometers",
            self.scale_um_var,
            "Example: type 20 if the visible scale bar is 20 µm.",
        )

        self._entry_row(
            parent,
            "Scale-bar length, pixels",
            self.scale_px_var,
            "Use Auto-detect bar pixels, or type the measured length manually.",
        )

        ttk.Button(parent, text="Auto-detect bar pixels", command=self.auto_detect_scale_pixels).pack(fill=tk.X, pady=(12, 5))
        ttk.Button(parent, text="Try OCR scale text", command=self.try_ocr_scale_text).pack(fill=tk.X, pady=(0, 5))
        ttk.Button(parent, text="Apply scale", command=self.apply_scale).pack(fill=tk.X, pady=(0, 5))
        ttk.Button(parent, text="Clear scale", command=self.clear_scale).pack(fill=tk.X)

        ttk.Label(parent, textvariable=self.pixel_size_label_var, wraplength=390, foreground="#333333").pack(fill=tk.X, pady=(12, 0))

        ttk.Label(
            parent,
            text=(
                "Recommended: auto-detect the pixel length, manually type the printed micrometer value, "
                "then apply scale."
            ),
            wraplength=390,
            foreground="#555555",
        ).pack(fill=tk.X, pady=(12, 0))

    def _build_presets_tab(self, parent):
        ttk.Label(
            parent,
            text=(
                "Save a named parameter set and reuse it for multiple image series. "
                "Presets include detection settings and scale calibration."
            ),
            wraplength=390,
        ).pack(fill=tk.X)

        self._entry_row(
            parent,
            "New preset name",
            self.preset_name_var,
            "Example: PNH_2W_oT_20x_dark_droplets",
            is_param=False,
        )

        ttk.Button(parent, text="Save current parameters as preset", command=self.save_preset).pack(fill=tk.X, pady=(10, 5))

        ttk.Label(parent, text="Saved presets").pack(anchor=tk.W, pady=(12, 0))
        self.preset_combo = ttk.Combobox(parent, textvariable=self.preset_select_var, state="readonly")
        self.preset_combo.pack(fill=tk.X)

        ttk.Button(parent, text="Load selected preset", command=self.load_selected_preset).pack(fill=tk.X, pady=(10, 5))
        ttk.Button(parent, text="Delete selected preset", command=self.delete_selected_preset).pack(fill=tk.X, pady=(0, 5))
        ttk.Button(parent, text="Refresh preset list", command=self.refresh_preset_list).pack(fill=tk.X)

        ttk.Label(
            parent,
            text=(
                "The currently loaded preset name is written to the Excel parameters sheet. "
                "If you manually change a parameter after loading, the name is still recorded, but the actual parameter values are also recorded."
            ),
            wraplength=390,
            foreground="#555555",
        ).pack(fill=tk.X, pady=(12, 0))

    def _build_export_tab(self, parent):
        ttk.Label(
            parent,
            text=(
                "Excel export analyzes all loaded frames using the current parameters."
            ),
            wraplength=390,
        ).pack(fill=tk.X)

        ttk.Button(parent, text="Export Excel", command=self.export_excel).pack(fill=tk.X, pady=(12, 0))

        ttk.Label(
            parent,
            text=(
                "Sheets: per_droplet, per_frame_summary, and parameters. "
                "The parameters sheet records every GUI parameter, scale calibration, loaded preset name, and software version."
            ),
            wraplength=390,
            foreground="#555555",
        ).pack(fill=tk.X, pady=(12, 0))

    def _build_help_tab(self, parent):
        help_text = """Overlay colors
Red circle: accepted droplet included in the count.
Green dot: droplet center, also called centroid.
Orange circle: rejected candidate, shown only if the diagnostic option is enabled.

V6 main controls
Droplet sensitivity: controls how easily weak/small candidates are found.
Shape strictness: the main false-positive protector. It raises roundness, circularity, edge-coverage, contrast, and blank-guard requirements together.
Close-droplet separation: controls splitting of neighboring droplets and how aggressively nearby detections are merged or kept separate.

V6 detector logic
The detector first finds dark local blob candidates, optionally splits large blobs into multiple local centers, then tests each candidate for local contrast, edge sharpness, roundness, circularity, eccentricity, and angular edge coverage.

Shape-first filters
A clean droplet should be compact and round, with boundary/contrast around most of the circumference. Non-round aggregates can therefore be rejected even if their raw intensity score is high.

Blank-frame guard
The blank guard now uses roundness protection. A convincing round droplet protects the frame from being blank-guarded, while frames containing only non-round debris can be suppressed.

Publication workflow
Tune the three Main sliders on representative frames, show rejected candidates to diagnose false positives, save the settings as a preset, lock parameters, export, and manually validate several frames from each condition."""

        ttk.Label(parent, text=help_text, wraplength=390, justify=tk.LEFT).pack(fill=tk.X)

    def _build_status_bar(self):
        bar = ttk.Frame(self.root, relief=tk.SUNKEN)
        bar.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Label(bar, textvariable=self.status_var, anchor=tk.W).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)

    def _bind_param_updates(self):
        for key, var in self.params.items():
            var.trace_add("write", lambda *args, key=key: self.on_parameter_changed(key))
        self.scale_um_var.trace_add("write", lambda *args: self.on_parameter_changed("scale_bar_um"))
        self.scale_px_var.trace_add("write", lambda *args: self.on_parameter_changed("scale_bar_px"))

    @staticmethod
    def _clip_float(value, low, high):
        try:
            value = float(value)
        except Exception:
            value = low
        return min(max(value, low), high)

    @staticmethod
    def estimate_candidate_roundness(
        patch_response: np.ndarray,
        patch_z: np.ndarray,
        inside: np.ndarray,
        annulus: np.ndarray,
        local_y: float,
        local_x: float,
        radius: float,
    ) -> Dict[str, float]:
        """
        Estimate how droplet-like/round a candidate is.

        The blank-frame guard uses this to avoid suppressing frames that contain a
        convincing round object. The score combines two independent ideas:
        1. connected-component shape circularity/solidity/eccentricity;
        2. angular coverage of the droplet contrast around the center.

        The returned roundness_score is 0 to 1. Higher means more round.
        """
        try:
            radius = max(float(radius), 1.0)
            py, px = patch_response.shape
            yy, xx = np.ogrid[:py, :px]
            d = np.sqrt((yy - local_y) ** 2 + (xx - local_x) ** 2)
            roi = d <= radius * 1.65

            if inside.sum() < 4 or roi.sum() < 8:
                return {
                    "roundness_score": 0.0,
                    "shape_circularity": 0.0,
                    "shape_solidity": 0.0,
                    "shape_eccentricity": 1.0,
                    "angular_coverage": 0.0,
                    "angular_uniformity": 0.0,
                }

            if annulus.any():
                baseline = float(np.median(patch_response[annulus]))
            else:
                baseline = float(np.percentile(patch_response[roi], 25))

            inner_signal = float(np.percentile(patch_response[inside], 75))
            contrast_span = max(inner_signal - baseline, 1e-6)
            threshold = baseline + 0.35 * contrast_span

            mask = (patch_response >= threshold) & roi
            mask = morphology.remove_small_objects(mask.astype(bool), min_size=4)

            shape_circularity = 0.0
            shape_solidity = 0.0
            shape_eccentricity = 1.0
            size_score = 0.0

            lab = measure.label(mask)
            if lab.max() > 0:
                cy = int(min(max(round(local_y), 0), py - 1))
                cx = int(min(max(round(local_x), 0), px - 1))
                center_label = int(lab[cy, cx])

                props = measure.regionprops(lab)
                chosen = None
                if center_label > 0:
                    for prop in props:
                        if prop.label == center_label:
                            chosen = prop
                            break

                if chosen is None and props:
                    # Fall back to the component whose centroid is closest to the LoG center.
                    chosen = min(
                        props,
                        key=lambda prop: math.hypot(prop.centroid[0] - local_y, prop.centroid[1] - local_x),
                    )

                if chosen is not None and chosen.area >= 4:
                    area = float(chosen.area)
                    perimeter = float(chosen.perimeter) if chosen.perimeter else 0.0
                    if perimeter > 0:
                        shape_circularity = 4.0 * math.pi * area / (perimeter * perimeter)
                    shape_circularity = DropletCounterApp._clip_float(shape_circularity, 0.0, 1.0)
                    shape_solidity = DropletCounterApp._clip_float(getattr(chosen, "solidity", 0.0), 0.0, 1.0)
                    shape_eccentricity = DropletCounterApp._clip_float(getattr(chosen, "eccentricity", 1.0), 0.0, 1.0)
                    eq_radius = math.sqrt(area / math.pi)
                    size_score = math.exp(-abs(eq_radius - radius) / max(radius * 0.75, 1.0))
                    size_score = DropletCounterApp._clip_float(size_score, 0.0, 1.0)

            shape_score = (
                0.42 * shape_circularity
                + 0.26 * (1.0 - shape_eccentricity)
                + 0.18 * shape_solidity
                + 0.14 * size_score
            )

            # Angular coverage: true droplets should produce contrast around most angles.
            theta = np.arctan2(yy - local_y, xx - local_x)
            n_sectors = 16
            sector_values = []
            for k in range(n_sectors):
                lo = -math.pi + 2.0 * math.pi * k / n_sectors
                hi = -math.pi + 2.0 * math.pi * (k + 1) / n_sectors
                sector = (theta >= lo) & (theta < hi)
                sector_inside = inside & sector
                sector_annulus = annulus & sector
                if sector_inside.sum() < 2:
                    continue
                if sector_annulus.sum() >= 2:
                    sector_baseline = float(np.median(patch_response[sector_annulus]))
                else:
                    sector_baseline = baseline
                sector_values.append(float(np.median(patch_response[sector_inside]) - sector_baseline))

            if sector_values:
                vals = np.asarray(sector_values, dtype=float)
                positive_cutoff = max(0.002, 0.20 * max(float(np.max(vals)), 1e-6))
                angular_coverage = float(np.mean(vals > positive_cutoff))
                positive_vals = vals[vals > positive_cutoff]
                if len(positive_vals) >= 2:
                    angular_uniformity = 1.0 - float(np.std(positive_vals) / (abs(np.mean(positive_vals)) + 1e-6))
                    angular_uniformity = DropletCounterApp._clip_float(angular_uniformity, 0.0, 1.0)
                else:
                    angular_uniformity = 0.0
            else:
                angular_coverage = 0.0
                angular_uniformity = 0.0

            angular_score = 0.72 * angular_coverage + 0.28 * angular_uniformity
            roundness_score = 0.55 * angular_score + 0.45 * shape_score
            roundness_score = DropletCounterApp._clip_float(roundness_score, 0.0, 1.0)

            return {
                "roundness_score": float(roundness_score),
                "shape_circularity": float(shape_circularity),
                "shape_solidity": float(shape_solidity),
                "shape_eccentricity": float(shape_eccentricity),
                "angular_coverage": float(angular_coverage),
                "angular_uniformity": float(angular_uniformity),
            }
        except Exception:
            return {
                "roundness_score": 0.0,
                "shape_circularity": 0.0,
                "shape_solidity": 0.0,
                "shape_eccentricity": 1.0,
                "angular_coverage": 0.0,
                "angular_uniformity": 0.0,
            }

    def sync_advanced_from_main(self, refresh: bool = True):
        """Map the three simple Main sliders onto the detailed Advanced settings."""
        if not bool(self.params["link_main_to_advanced"].get()):
            self.update_effective_thresholds_label()
            return

        sensitivity = self._clip_float(self.params["droplet_sensitivity"].get(), 1.0, 100.0)
        shape = self._clip_float(self.params["droplet_shape_strictness"].get(), 1.0, 100.0)
        separation = self._clip_float(self.params["close_droplet_separation"].get(), 1.0, 100.0)

        s_norm = (sensitivity - 1.0) / 99.0
        sh_norm = (shape - 1.0) / 99.0
        sep_norm = (separation - 1.0) / 99.0

        # Candidate discovery. Lower threshold means more raw candidates.
        candidate_z_threshold = 3.85 - 3.00 * s_norm

        # Final intensity gates. Shape strictness tightens; sensitivity mildly relaxes.
        quality_threshold = 1.05 + 2.45 * sh_norm - 0.40 * s_norm
        min_contrast_required = 0.008 + 0.050 * sh_norm - 0.010 * s_norm
        min_edge_required = 0.006 + 0.046 * sh_norm - 0.008 * s_norm

        # Shape-first filters. These are the main non-droplet/blob rejection gates.
        min_roundness_required = 0.30 + 0.48 * sh_norm
        min_shape_circularity_required = 0.16 + 0.46 * sh_norm
        max_shape_eccentricity_allowed = 0.94 - 0.42 * sh_norm
        min_angular_coverage_required = 0.24 + 0.48 * sh_norm
        min_angular_uniformity_required = 0.02 + 0.32 * sh_norm

        # Close-droplet separation. Higher separation preserves nearby candidates and splits large blobs.
        signal_blur_sigma = 0.92 - 0.72 * sep_norm
        max_overlap_fraction = 0.25 + 0.55 * sep_norm
        nms_overlap_factor = 1.10 - 0.72 * sep_norm
        split_peak_threshold = 3.00 - 1.70 * sep_norm
        split_peak_distance_fraction = 0.76 - 0.42 * sep_norm
        split_child_radius_fraction = 0.60 - 0.14 * sep_norm

        # Blank-frame guard. A round droplet should protect the frame.
        blank_frame_score_threshold = 2.85 + 1.65 * sh_norm - 0.25 * s_norm
        blank_guard_roundness_threshold = max(0.36, min_roundness_required - 0.05)

        mapped = {
            "false_positive_rejection": shape,
            "candidate_z_threshold": self._clip_float(candidate_z_threshold, 0.75, 4.0),
            "quality_threshold": self._clip_float(quality_threshold, 0.8, 4.2),
            "min_contrast_required": self._clip_float(min_contrast_required, 0.0, 0.12),
            "min_edge_required": self._clip_float(min_edge_required, 0.0, 0.12),
            "min_roundness_required": self._clip_float(min_roundness_required, 0.20, 0.95),
            "min_shape_circularity_required": self._clip_float(min_shape_circularity_required, 0.05, 0.95),
            "max_shape_eccentricity_allowed": self._clip_float(max_shape_eccentricity_allowed, 0.20, 0.98),
            "min_angular_coverage_required": self._clip_float(min_angular_coverage_required, 0.05, 1.0),
            "min_angular_uniformity_required": self._clip_float(min_angular_uniformity_required, 0.0, 0.90),
            "signal_blur_sigma": self._clip_float(signal_blur_sigma, 0.0, 3.0),
            "max_overlap_fraction": self._clip_float(max_overlap_fraction, 0.1, 0.9),
            "nms_overlap_factor": self._clip_float(nms_overlap_factor, 0.2, 1.5),
            "split_peak_threshold": self._clip_float(split_peak_threshold, 0.50, 5.00),
            "split_peak_distance_fraction": self._clip_float(split_peak_distance_fraction, 0.20, 1.00),
            "split_child_radius_fraction": self._clip_float(split_child_radius_fraction, 0.25, 0.80),
            "blank_frame_score_threshold": self._clip_float(blank_frame_score_threshold, 2.0, 7.0),
            "blank_guard_roundness_threshold": self._clip_float(blank_guard_roundness_threshold, 0.30, 0.95),
        }

        self._syncing_params = True
        try:
            for key, value in mapped.items():
                self.params[key].set(round(value, 4))
        finally:
            self._syncing_params = False

        self.update_effective_thresholds_label()
        if refresh and not self.parameters_locked_var.get():
            self.refresh_preview(debounce=True)

    def update_effective_thresholds_label(self):
        try:
            txt = (
                f"Effective: candidate Z ≥ {float(self.params['candidate_z_threshold'].get()):.2f}; "
                f"score ≥ {float(self.params['quality_threshold'].get()):.2f}; "
                f"roundness ≥ {float(self.params['min_roundness_required'].get()):.2f}; "
                f"edge coverage ≥ {float(self.params['min_angular_coverage_required'].get()):.2f}; "
                f"NMS distance {float(self.params['nms_overlap_factor'].get()):.2f}."
            )
        except Exception:
            txt = "Effective thresholds: unavailable."
        self.effective_thresholds_var.set(txt)

    def on_parameter_changed(self, changed_key=None):
        if self._syncing_params:
            return

        display_only_keys = {"preview_gamma", "show_labels", "show_rejected_candidates"}

        if changed_key in {"droplet_sensitivity", "droplet_shape_strictness", "close_droplet_separation", "false_positive_rejection", "link_main_to_advanced"}:
            self.sync_advanced_from_main(refresh=False)
            if changed_key == "link_main_to_advanced":
                self.update_lock_state()
        else:
            self.update_effective_thresholds_label()

        if self.parameters_locked_var.get():
            return

        if changed_key in display_only_keys and self.current_detection is not None:
            self.render_current_detection_to_canvas()
        else:
            self.refresh_preview(debounce=True)

    def update_lock_state(self):
        locked = self.parameters_locked_var.get()
        state = "disabled" if locked else "normal"

        for widget in self.parameter_widgets:
            try:
                widget.configure(state=state)
            except Exception:
                pass

        if not locked and bool(self.params["link_main_to_advanced"].get()):
            for widget in self.linked_advanced_widgets:
                try:
                    widget.configure(state="disabled")
                except Exception:
                    pass

        self.update_effective_thresholds_label()

        if locked:
            self.status_var.set("Parameters locked.")
        else:
            self.status_var.set("Parameters unlocked.")
            self.refresh_preview(debounce=True)

    def load_folder(self):
        folder = filedialog.askdirectory(title="Select folder with image series")
        if not folder:
            return

        paths = [p for p in Path(folder).iterdir() if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS]
        paths.sort(key=natural_sort_key)

        if not paths:
            messagebox.showwarning("No images", "No supported image files were found in this folder.")
            return

        new_frames: List[FrameRecord] = []
        failed = []
        for p in paths:
            try:
                new_frames.extend(load_image_file(p))
            except Exception:
                failed.append(p.name)

        if not new_frames:
            messagebox.showerror("Load failed", "No images could be loaded.")
            return

        self.frames = new_frames
        self.clear_caches()
        self.current_index = 0
        self.frame_slider.configure(from_=0, to=max(0, len(self.frames) - 1))
        self.frame_slider.set(0)
        self.status_var.set(f"Loaded {len(self.frames)} frame(s) from folder. Failed files: {len(failed)}.")
        self.refresh_preview(debounce=False)

    def load_single_file(self):
        filetypes = [
            ("Image files", "*.tif *.tiff *.png *.jpg *.jpeg *.bmp"),
            ("TIFF files", "*.tif *.tiff"),
            ("All files", "*.*"),
        ]
        filename = filedialog.askopenfilename(title="Select image or TIFF stack", filetypes=filetypes)
        if not filename:
            return

        try:
            new_frames = load_image_file(Path(filename))
        except Exception as exc:
            messagebox.showerror("Load failed", f"Could not load file:\n{exc}")
            return

        if not new_frames:
            messagebox.showerror("Load failed", "No image frames were found.")
            return

        self.frames = new_frames
        self.clear_caches()
        self.current_index = 0
        self.frame_slider.configure(from_=0, to=max(0, len(self.frames) - 1))
        self.frame_slider.set(0)
        self.status_var.set(f"Loaded {len(self.frames)} frame(s) from single file.")
        self.refresh_preview(debounce=False)

    def prev_frame(self):
        if not self.frames:
            return
        self.current_index = max(0, self.current_index - 1)
        self.frame_slider.set(self.current_index)
        self.refresh_preview(debounce=False)

    def next_frame(self):
        if not self.frames:
            return
        self.current_index = min(len(self.frames) - 1, self.current_index + 1)
        self.frame_slider.set(self.current_index)
        self.refresh_preview(debounce=False)

    def on_slider_change(self, value):
        if not self.frames:
            return
        try:
            idx = int(round(float(value)))
        except Exception:
            return
        idx = max(0, min(len(self.frames) - 1, idx))
        if idx != self.current_index:
            self.current_index = idx
            self.refresh_preview(debounce=True)

    def get_current_frame(self) -> Optional[FrameRecord]:
        if not self.frames:
            return None
        return self.frames[self.current_index]

    def get_pixel_size_um(self) -> Optional[float]:
        try:
            um = float(self.scale_um_var.get())
            px = float(self.scale_px_var.get())
            if um > 0 and px > 0:
                return um / px
        except Exception:
            pass
        return None

    def update_pixel_size_label(self):
        pix = self.get_pixel_size_um()
        if pix is None:
            self.pixel_size_label_var.set("Pixel size: not calibrated")
        else:
            self.pixel_size_label_var.set(f"Pixel size: {pix:.6g} µm/pixel. Area scale: {pix * pix:.6g} µm²/pixel².")

    def apply_scale(self):
        pix = self.get_pixel_size_um()
        if pix is None:
            messagebox.showwarning("Scale not valid", "Enter positive values for both scale-bar micrometers and scale-bar pixels.")
            return
        self.update_pixel_size_label()
        self.refresh_preview(debounce=False)
        self.status_var.set(f"Scale applied: {pix:.6g} µm/pixel.")

    def clear_scale(self):
        self.scale_um_var.set("")
        self.scale_px_var.set("")
        self.update_pixel_size_label()
        self.refresh_preview(debounce=False)
        self.status_var.set("Scale cleared.")

    def auto_detect_scale_pixels(self):
        frame = self.get_current_frame()
        if frame is None:
            messagebox.showwarning("No image", "Load an image first.")
            return
        px = detect_scale_bar_pixels(frame.image)
        if px is None:
            messagebox.showwarning("Scale bar not found", "Could not detect a plausible horizontal scale bar. Type the pixel length manually.")
            return
        self.scale_px_var.set(f"{px:.2f}")
        self.update_pixel_size_label()
        self.status_var.set(f"Detected scale-bar length: {px:.2f} pixels.")

    def try_ocr_scale_text(self):
        frame = self.get_current_frame()
        if frame is None:
            messagebox.showwarning("No image", "Load an image first.")
            return

        val = ocr_scale_text(frame.image)
        if val is None:
            messagebox.showwarning(
                "OCR failed",
                "Could not read the scale text. Manual entry is usually more reliable."
            )
            return

        self.scale_um_var.set(f"{val:g}")
        self.update_pixel_size_label()
        self.status_var.set(f"OCR read scale text as {val:g} µm.")

    def get_params_values(self) -> Dict[str, object]:
        return {
            "software_version": SOFTWARE_VERSION,
            "current_preset_name": self.current_preset_name,
            "droplet_sensitivity": float(self.params["droplet_sensitivity"].get()),
            "droplet_shape_strictness": float(self.params["droplet_shape_strictness"].get()),
            "close_droplet_separation": float(self.params["close_droplet_separation"].get()),
            "false_positive_rejection": float(self.params["false_positive_rejection"].get()),
            "link_main_to_advanced": bool(self.params["link_main_to_advanced"].get()),
            "candidate_z_threshold": float(self.params["candidate_z_threshold"].get()),
            "quality_threshold": float(self.params["quality_threshold"].get()),
            "min_contrast_required": float(self.params["min_contrast_required"].get()),
            "min_edge_required": float(self.params["min_edge_required"].get()),
            "use_shape_filters": bool(self.params["use_shape_filters"].get()),
            "min_roundness_required": float(self.params["min_roundness_required"].get()),
            "min_shape_circularity_required": float(self.params["min_shape_circularity_required"].get()),
            "max_shape_eccentricity_allowed": float(self.params["max_shape_eccentricity_allowed"].get()),
            "min_angular_coverage_required": float(self.params["min_angular_coverage_required"].get()),
            "min_angular_uniformity_required": float(self.params["min_angular_uniformity_required"].get()),
            "split_close_droplets": bool(self.params["split_close_droplets"].get()),
            "split_min_diameter_px": float(self.params["split_min_diameter_px"].get()),
            "split_peak_threshold": float(self.params["split_peak_threshold"].get()),
            "split_peak_distance_fraction": float(self.params["split_peak_distance_fraction"].get()),
            "split_child_radius_fraction": float(self.params["split_child_radius_fraction"].get()),
            "split_max_children": int(round(float(self.params["split_max_children"].get()))),
            "bright_objects": bool(self.params["bright_objects"].get()),
            "preview_gamma": float(self.params["preview_gamma"].get()),
            "min_diameter_px": float(self.params["min_diameter_px"].get()),
            "max_diameter_px": float(self.params["max_diameter_px"].get()),
            "background_sigma": float(self.params["background_sigma"].get()),
            "signal_blur_sigma": float(self.params["signal_blur_sigma"].get()),
            "max_overlap_fraction": float(self.params["max_overlap_fraction"].get()),
            "nms_overlap_factor": float(self.params["nms_overlap_factor"].get()),
            "blank_frame_guard": bool(self.params["blank_frame_guard"].get()),
            "blank_frame_score_threshold": float(self.params["blank_frame_score_threshold"].get()),
            "blank_guard_roundness_threshold": float(self.params["blank_guard_roundness_threshold"].get()),
            "blank_guard_min_round_droplets": int(round(float(self.params["blank_guard_min_round_droplets"].get()))),
            "blank_guard_use_roundness": bool(self.params["blank_guard_use_roundness"].get()),
            "exclude_border": bool(self.params["exclude_border"].get()),
            "show_labels": bool(self.params["show_labels"].get()),
            "show_rejected_candidates": bool(self.params["show_rejected_candidates"].get()),
            "preview_markers_visible": bool(self.preview_markers_visible_var.get()),
            "scale_bar_um": self.scale_um_var.get(),
            "scale_bar_px": self.scale_px_var.get(),
            "pixel_size_um": self.get_pixel_size_um(),
            "parameters_locked": bool(self.parameters_locked_var.get()),
        }

    def set_params_values(self, values: Dict[str, object]):
        # Temporarily unlock while applying.
        was_locked = self.parameters_locked_var.get()
        self.parameters_locked_var.set(False)

        for key, var in self.params.items():
            if key not in values:
                continue

            val = values[key]
            try:
                if isinstance(var, tk.BooleanVar):
                    var.set(bool(val))
                elif isinstance(var, tk.DoubleVar):
                    var.set(float(val))
                elif isinstance(var, tk.StringVar):
                    var.set(str(val))
                else:
                    var.set(val)
            except Exception:
                pass

        if "scale_bar_um" in values:
            self.scale_um_var.set(str(values.get("scale_bar_um") or ""))
        if "scale_bar_px" in values:
            self.scale_px_var.set(str(values.get("scale_bar_px") or ""))

        # Backward compatibility: older presets had false_positive_rejection but not V6 shape strictness.
        if "droplet_shape_strictness" not in values and "false_positive_rejection" in values:
            try:
                self.params["droplet_shape_strictness"].set(float(values["false_positive_rejection"]))
            except Exception:
                pass

        self.parameters_locked_var.set(was_locked)
        if bool(self.params["link_main_to_advanced"].get()):
            self.sync_advanced_from_main(refresh=False)
        self.update_pixel_size_label()
        self.update_lock_state()
        self.refresh_preview(debounce=False)

    def read_presets(self) -> Dict[str, Dict[str, object]]:
        if not self.preset_file.exists():
            return {}
        try:
            with open(self.preset_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
        except Exception:
            pass
        return {}

    def write_presets(self, presets: Dict[str, Dict[str, object]]):
        with open(self.preset_file, "w", encoding="utf-8") as f:
            json.dump(presets, f, indent=2)

    def refresh_preset_list(self):
        presets = self.read_presets()
        names = sorted(presets.keys())
        try:
            self.preset_combo.configure(values=names)
        except Exception:
            pass
        if names and self.preset_select_var.get() not in names:
            self.preset_select_var.set(names[0])
        elif not names:
            self.preset_select_var.set("")

    def save_preset(self):
        name = self.preset_name_var.get().strip()
        if not name:
            messagebox.showwarning("Preset name missing", "Type a name for the preset.")
            return

        presets = self.read_presets()
        values = self.get_params_values()
        values["current_preset_name"] = name
        values["software_version"] = SOFTWARE_VERSION
        presets[name] = values
        self.write_presets(presets)

        self.current_preset_name = name
        self.refresh_preset_list()
        self.preset_select_var.set(name)
        self.status_var.set(f"Saved preset: {name}")

    def load_selected_preset(self):
        name = self.preset_select_var.get().strip()
        if not name:
            messagebox.showwarning("No preset selected", "Select a preset first.")
            return

        presets = self.read_presets()
        if name not in presets:
            messagebox.showwarning("Preset not found", "The selected preset was not found.")
            self.refresh_preset_list()
            return

        self.current_preset_name = name
        self.set_params_values(presets[name])
        self.status_var.set(f"Loaded preset: {name}")

    def delete_selected_preset(self):
        name = self.preset_select_var.get().strip()
        if not name:
            return

        if not messagebox.askyesno("Delete preset", f"Delete preset '{name}'?"):
            return

        presets = self.read_presets()
        if name in presets:
            del presets[name]
            self.write_presets(presets)

        if self.current_preset_name == name:
            self.current_preset_name = None

        self.refresh_preset_list()
        self.status_var.set(f"Deleted preset: {name}")

    def clear_caches(self):
        self.signal_cache.clear()
        self.analysis_cache.clear()
        self.current_detection = None
        self.current_detection_cache_key = None

    @staticmethod
    def _rounded_for_key(value):
        if isinstance(value, float):
            return round(value, 6)
        return value

    def detection_cache_key(self, frame_index: int, frame_name: str, params: Dict[str, object]) -> tuple:
        """Parameters that change detection/measurement results. Display-only settings are excluded."""
        keys = [
            "droplet_sensitivity", "droplet_shape_strictness", "close_droplet_separation",
            "false_positive_rejection", "link_main_to_advanced",
            "candidate_z_threshold", "quality_threshold", "min_contrast_required", "min_edge_required",
            "use_shape_filters", "min_roundness_required", "min_shape_circularity_required",
            "max_shape_eccentricity_allowed", "min_angular_coverage_required", "min_angular_uniformity_required",
            "split_close_droplets", "split_min_diameter_px", "split_peak_threshold",
            "split_peak_distance_fraction", "split_child_radius_fraction", "split_max_children",
            "bright_objects", "min_diameter_px", "max_diameter_px", "background_sigma",
            "signal_blur_sigma", "max_overlap_fraction", "nms_overlap_factor",
            "blank_frame_guard", "blank_frame_score_threshold", "blank_guard_roundness_threshold",
            "blank_guard_min_round_droplets", "blank_guard_use_roundness", "exclude_border",
            "pixel_size_um",
        ]
        return (frame_index, frame_name, tuple((k, self._rounded_for_key(params.get(k))) for k in keys))

    def signal_cache_key(self, frame_index: int, frame_name: str, params: Dict[str, object]) -> tuple:
        return (
            frame_index,
            frame_name,
            bool(params.get("bright_objects")),
            round(float(params.get("background_sigma", 18.0)), 6),
            round(float(params.get("signal_blur_sigma", 0.45)), 6),
        )

    def _cache_put_limited(self, cache: Dict, key, value, max_entries: int):
        cache[key] = value
        while len(cache) > max_entries:
            oldest = next(iter(cache))
            cache.pop(oldest, None)

    def get_cached_signal(self, image: np.ndarray, frame_index: int, frame_name: str, params: Dict[str, object]):
        key = self.signal_cache_key(frame_index, frame_name, params)
        cached = self.signal_cache.get(key)
        if cached is not None:
            return cached
        value = self.compute_signal(image, params)
        self._cache_put_limited(self.signal_cache, key, value, self.max_signal_cache_entries)
        return value

    def get_or_analyze_detection(self, frame: FrameRecord) -> DetectionOutput:
        params = self.get_params_values()
        key = self.detection_cache_key(self.current_index, frame.name, params)
        cached = self.analysis_cache.get(key)
        if cached is None:
            accepted, rejected, metrics = self.detect_candidates(
                frame.image,
                frame_index=self.current_index,
                frame_name=frame.name,
                params_override=params,
            )
            cached = (accepted, rejected, metrics)
            self._cache_put_limited(self.analysis_cache, key, cached, self.max_analysis_cache_entries)
        accepted, rejected, metrics = cached
        output = DetectionOutput(None, accepted, rejected, len(accepted), metrics)
        self.current_detection = output
        self.current_detection_cache_key = key
        return output

    def refresh_preview(self, debounce: bool = True):
        if debounce:
            if self.update_after_id is not None:
                self.root.after_cancel(self.update_after_id)
            # V6 shape analysis is expensive; a longer debounce prevents re-analysis
            # for every tiny slider movement while dragging.
            self.update_after_id = self.root.after(350, self._refresh_preview_now)
        else:
            if self.update_after_id is not None:
                self.root.after_cancel(self.update_after_id)
                self.update_after_id = None
            self._refresh_preview_now()

    def _refresh_preview_now(self):
        self.update_after_id = None
        frame = self.get_current_frame()

        if frame is None:
            self.canvas.delete("all")
            self.frame_label_var.set("No image loaded")
            return

        try:
            self.status_var.set("Analyzing current frame...")
            self.root.update_idletasks()
            output = self.get_or_analyze_detection(frame)
            self.render_detection_to_canvas(output)
            self.status_var.set("Preview updated.")
        except Exception as exc:
            self.canvas.delete("all")
            self.frame_label_var.set(f"Frame {self.current_index + 1}/{len(self.frames)}: analysis failed")
            self.status_var.set(f"Analysis error: {exc}")

    def render_current_detection_to_canvas(self):
        frame = self.get_current_frame()
        if frame is None:
            return

        try:
            expected_key = self.detection_cache_key(self.current_index, frame.name, self.get_params_values())
        except Exception:
            expected_key = None

        if self.current_detection is not None and self.current_detection_cache_key == expected_key:
            self.render_detection_to_canvas(self.current_detection)
        else:
            self.refresh_preview(debounce=False)

    def render_detection_to_canvas(self, output: DetectionOutput):
        frame = self.get_current_frame()
        if frame is None:
            return

        params = self.get_params_values()
        preview_gray = display_to_float01(frame.image, gamma=float(params.get("preview_gamma", 0.75)))
        rejected_for_preview = output.rejected_rows if bool(params.get("show_rejected_candidates")) else []
        overlay = self.make_overlay(
            preview_gray,
            output.rows,
            rejected_for_preview,
            bool(params.get("show_labels")),
            show_markers=bool(self.preview_markers_visible_var.get()),
        )
        output.overlay_image = overlay
        display = self.fit_image_to_canvas(overlay)
        self.preview_photo = ImageTk.PhotoImage(display)
        self.canvas.delete("all")
        self.canvas.create_image(
            self.canvas.winfo_width() // 2,
            self.canvas.winfo_height() // 2,
            image=self.preview_photo,
            anchor=tk.CENTER,
        )

        guard_txt = " | blank guard applied" if output.frame_metrics.get("blank_frame_guard_applied") else ""
        self.frame_label_var.set(
            f"Frame {self.current_index + 1}/{len(self.frames)}: {frame.name} | droplets: {output.count}{guard_txt}"
        )

    def fit_image_to_canvas(self, pil_img: Image.Image) -> Image.Image:
        cw = max(1, self.canvas.winfo_width())
        ch = max(1, self.canvas.winfo_height())
        iw, ih = pil_img.size
        scale = min(cw / iw, ch / ih)
        new_w = max(1, int(iw * scale))
        new_h = max(1, int(ih * scale))
        return pil_img.resize((new_w, new_h), Image.Resampling.LANCZOS)

    def compute_signal(self, image: np.ndarray, params: Dict[str, object]) -> Tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray]:
        gray = robust_normalize(image)
        bg_sigma = max(1.0, float(params["background_sigma"]))
        background = ndi.gaussian_filter(gray, sigma=bg_sigma)

        if bool(params["bright_objects"]):
            response = np.clip(gray - background, 0, None)
        else:
            response = np.clip(background - gray, 0, None)

        blur = float(params["signal_blur_sigma"])
        if blur > 0:
            response = ndi.gaussian_filter(response, sigma=blur)

        med = float(np.median(response))
        mad = float(np.median(np.abs(response - med))) + 1e-6
        z_signal = np.clip((response - med) / (1.4826 * mad), 0, None)

        edge_image = filters.sobel(gray)
        return gray, response, z_signal, edge_image

    def expand_blobs_with_splitting(
        self,
        blobs: np.ndarray,
        z_signal: np.ndarray,
        params: Dict[str, object],
        min_sigma: float,
        max_sigma: float,
    ) -> Tuple[np.ndarray, Dict[str, int]]:
        """Split large LoG candidates into multiple local centers when close droplets are merged."""
        metrics = {
            "original_blob_candidates": int(len(blobs)),
            "split_parent_candidates": 0,
            "split_child_candidates": 0,
            "candidate_blobs_after_splitting": int(len(blobs)),
        }

        if not bool(params.get("split_close_droplets", True)) or blobs is None or len(blobs) == 0:
            return blobs, metrics

        h, w = z_signal.shape
        split_min_d = float(params.get("split_min_diameter_px", 8.0))
        peak_threshold = float(params.get("split_peak_threshold", 2.0))
        distance_fraction = self._clip_float(params.get("split_peak_distance_fraction", 0.45), 0.20, 1.00)
        child_radius_fraction = self._clip_float(params.get("split_child_radius_fraction", 0.50), 0.25, 0.80)
        max_children = int(round(self._clip_float(params.get("split_max_children", 6), 2, 12)))

        expanded: List[List[float]] = []
        for blob in np.asarray(blobs, dtype=float):
            y, x, sigma = float(blob[0]), float(blob[1]), float(blob[2])
            radius = float(sigma * math.sqrt(2.0))
            diameter_px = 2.0 * radius

            if diameter_px < split_min_d:
                expanded.append([y, x, sigma])
                continue

            y0 = int(max(0, y - radius * 1.6))
            y1 = int(min(h, y + radius * 1.6 + 1))
            x0 = int(max(0, x - radius * 1.6))
            x1 = int(min(w, x + radius * 1.6 + 1))
            if y1 <= y0 or x1 <= x0:
                expanded.append([y, x, sigma])
                continue

            patch = z_signal[y0:y1, x0:x1]
            yy, xx = np.ogrid[y0:y1, x0:x1]
            roi = ((yy - y) ** 2 + (xx - x) ** 2) <= (radius * 1.25) ** 2
            min_distance = max(2, int(round(radius * distance_fraction)))

            try:
                peaks = feature.peak_local_max(
                    patch,
                    min_distance=min_distance,
                    threshold_abs=peak_threshold,
                    labels=roi.astype(np.uint8),
                    exclude_border=False,
                    num_peaks=max_children,
                )
            except Exception:
                peaks = np.empty((0, 2), dtype=int)

            if len(peaks) >= 2:
                metrics["split_parent_candidates"] += 1
                child_radius = self._clip_float(radius * child_radius_fraction, 0.8 * math.sqrt(2.0), radius)
                child_sigma = self._clip_float(child_radius / math.sqrt(2.0), min_sigma, max_sigma)
                for py, px in peaks:
                    expanded.append([float(y0 + py), float(x0 + px), float(child_sigma)])
                metrics["split_child_candidates"] += int(len(peaks))
            else:
                expanded.append([y, x, sigma])

        out = np.asarray(expanded, dtype=float) if expanded else np.empty((0, 3), dtype=float)
        metrics["candidate_blobs_after_splitting"] = int(len(out))
        return out, metrics

    def detect_candidates(self, image: np.ndarray, frame_index: int, frame_name: str, params_override: Optional[Dict[str, object]] = None) -> Tuple[List[Dict[str, object]], List[Dict[str, object]], Dict[str, object]]:
        params = params_override if params_override is not None else self.get_params_values()
        temperature_c = extract_temperature_c_from_name(frame_name)
        gray, response, z_signal, edge_image = self.get_cached_signal(image, frame_index, frame_name, params)

        sensitivity = float(params["droplet_sensitivity"])
        rejection = float(params["false_positive_rejection"])
        min_d = max(1.5, float(params["min_diameter_px"]))
        max_d = max(min_d + 1.0, float(params["max_diameter_px"]))
        pixel_size_um = params["pixel_size_um"]

        min_sigma = max(0.8, (min_d / 2.0) / math.sqrt(2.0))
        max_sigma = max(min_sigma + 0.1, (max_d / 2.0) / math.sqrt(2.0))

        # In V4.1 the Main sliders can drive these effective Advanced thresholds.
        # When linked mode is off, the values are read directly from the Advanced tab.
        candidate_z_threshold = self._clip_float(params.get("candidate_z_threshold", 3.6 - 2.6 * (sensitivity / 100.0)), 0.75, 4.0)

        try:
            raw_blobs = feature.blob_log(
                z_signal,
                min_sigma=min_sigma,
                max_sigma=max_sigma,
                num_sigma=10,
                threshold=candidate_z_threshold,
                overlap=float(params["max_overlap_fraction"]),
            )
        except Exception:
            raw_blobs = np.empty((0, 3), dtype=float)

        blobs, split_metrics = self.expand_blobs_with_splitting(
            raw_blobs,
            z_signal=z_signal,
            params=params,
            min_sigma=min_sigma,
            max_sigma=max_sigma,
        )

        h, w = gray.shape
        accepted_before_nms: List[Dict[str, object]] = []
        rejected: List[Dict[str, object]] = []

        q_threshold = self._clip_float(params.get("quality_threshold", 1.4 + (rejection / 100.0) * 2.0), 0.8, 4.2)
        min_contrast = self._clip_float(params.get("min_contrast_required", 0.015 + (rejection / 100.0) * 0.035), 0.0, 0.12)
        min_edge = self._clip_float(params.get("min_edge_required", 0.010 + (rejection / 100.0) * 0.035), 0.0, 0.12)

        for blob in blobs:
            y, x, sigma = float(blob[0]), float(blob[1]), float(blob[2])
            radius = float(sigma * math.sqrt(2.0))
            diameter_px = 2.0 * radius

            reason = ""
            if diameter_px < min_d:
                reason = "below_min_diameter"
            elif diameter_px > max_d:
                reason = "above_max_diameter"

            if not reason and bool(params["exclude_border"]):
                if y < radius * 1.2 or x < radius * 1.2 or y >= h - radius * 1.2 or x >= w - radius * 1.2:
                    reason = "touches_border"

            y0 = int(max(0, y - radius * 3.0))
            y1 = int(min(h, y + radius * 3.0 + 1))
            x0 = int(max(0, x - radius * 3.0))
            x1 = int(min(w, x + radius * 3.0 + 1))

            if y1 <= y0 or x1 <= x0:
                continue

            yy, xx = np.ogrid[y0:y1, x0:x1]
            d = np.sqrt((yy - y) ** 2 + (xx - x) ** 2)

            inside = d <= radius
            annulus = (d >= radius * 1.4) & (d <= radius * 2.8)
            ring = (d >= radius * 0.8) & (d <= radius * 1.5)

            if inside.sum() < 4 or annulus.sum() < 8:
                continue

            patch_gray = gray[y0:y1, x0:x1]
            patch_response = response[y0:y1, x0:x1]
            patch_z = z_signal[y0:y1, x0:x1]
            patch_edge = edge_image[y0:y1, x0:x1]
            roundness_metrics = {
                "roundness_score": 0.0,
                "shape_circularity": 0.0,
                "shape_solidity": 0.0,
                "shape_eccentricity": 1.0,
                "angular_coverage": 0.0,
                "angular_uniformity": 0.0,
            }

            if bool(params["bright_objects"]):
                local_contrast = float(np.median(patch_gray[inside]) - np.median(patch_gray[annulus]))
            else:
                local_contrast = float(np.median(patch_gray[annulus]) - np.median(patch_gray[inside]))

            local_edge = float(np.mean(patch_edge[ring])) if ring.any() else 0.0
            signal_z = float(np.median(patch_z[inside]))

            detection_score = local_contrast * 4.0 + local_edge * 8.0 + min(signal_z / 8.0, 1.0)

            if not reason:
                if detection_score < q_threshold:
                    reason = "low_detection_score"
                elif local_contrast < min_contrast:
                    reason = "low_local_contrast"
                elif local_edge < min_edge:
                    reason = "low_edge_sharpness"

            if not reason and (bool(params.get("use_shape_filters", True)) or bool(params.get("blank_frame_guard", True))):
                roundness_metrics = self.estimate_candidate_roundness(
                    patch_response=patch_response,
                    patch_z=patch_z,
                    inside=inside,
                    annulus=annulus,
                    local_y=y - y0,
                    local_x=x - x0,
                    radius=radius,
                )

            if not reason and bool(params.get("use_shape_filters", True)):
                if float(roundness_metrics["roundness_score"]) < float(params.get("min_roundness_required", 0.60)):
                    reason = "low_roundness"
                elif float(roundness_metrics["shape_circularity"]) < float(params.get("min_shape_circularity_required", 0.40)):
                    reason = "low_shape_circularity"
                elif float(roundness_metrics["shape_eccentricity"]) > float(params.get("max_shape_eccentricity_allowed", 0.80)):
                    reason = "high_shape_eccentricity"
                elif float(roundness_metrics["angular_coverage"]) < float(params.get("min_angular_coverage_required", 0.45)):
                    reason = "low_angular_edge_coverage"
                elif float(roundness_metrics["angular_uniformity"]) < float(params.get("min_angular_uniformity_required", 0.0)):
                    reason = "low_angular_uniformity"

            area_px2 = math.pi * radius * radius
            perimeter_px = 2.0 * math.pi * radius

            row = {
                "frame_index": frame_index + 1,
                "frame_name": frame_name,
                "temperature_C": temperature_c,
                "droplet_number": None,
                "centroid_x_px": x,
                "centroid_y_px": y,
                "radius_px": radius,
                "diameter_px": diameter_px,
                "area_px2": area_px2,
                "perimeter_px": perimeter_px,
                "local_contrast": local_contrast,
                "edge_sharpness": local_edge,
                "signal_z": signal_z,
                "detection_score": detection_score,
                "roundness_score": roundness_metrics["roundness_score"],
                "shape_circularity": roundness_metrics["shape_circularity"],
                "shape_solidity": roundness_metrics["shape_solidity"],
                "shape_eccentricity": roundness_metrics["shape_eccentricity"],
                "angular_coverage": roundness_metrics["angular_coverage"],
                "angular_uniformity": roundness_metrics["angular_uniformity"],
                "accepted": reason == "",
                "rejection_reason": reason,
            }

            if pixel_size_um is not None:
                row["radius_um"] = radius * pixel_size_um
                row["diameter_um"] = diameter_px * pixel_size_um
                row["area_um2"] = area_px2 * pixel_size_um * pixel_size_um
                row["perimeter_um"] = perimeter_px * pixel_size_um
            else:
                row["radius_um"] = None
                row["diameter_um"] = None
                row["area_um2"] = None
                row["perimeter_um"] = None

            if reason:
                rejected.append(row)
            else:
                accepted_before_nms.append(row)

        scores = [float(row["detection_score"]) for row in accepted_before_nms]
        roundness_scores = [float(row.get("roundness_score", 0.0)) for row in accepted_before_nms]
        max_score = max(scores) if scores else 0.0
        top5_mean = float(np.mean(sorted(scores, reverse=True)[:5])) if scores else 0.0
        max_roundness = max(roundness_scores) if roundness_scores else 0.0
        top5_roundness_mean = float(np.mean(sorted(roundness_scores, reverse=True)[:5])) if roundness_scores else 0.0

        guard_applied = False
        guard_reason = "not_applied"
        roundness_threshold = self._clip_float(params.get("blank_guard_roundness_threshold", 0.62), 0.30, 0.95)
        min_round_droplets = int(round(self._clip_float(params.get("blank_guard_min_round_droplets", 1), 1, 10)))
        round_protectors = [row for row in accepted_before_nms if float(row.get("roundness_score", 0.0)) >= roundness_threshold]
        round_protector_count = len(round_protectors)

        if bool(params["blank_frame_guard"]):
            guard_threshold = float(params["blank_frame_score_threshold"])
            weak_score_frame = max_score < guard_threshold and top5_mean < guard_threshold * 0.95
            no_round_droplet_frame = round_protector_count < min_round_droplets

            if bool(params.get("blank_guard_use_roundness", True)):
                # New behavior: a convincing round droplet protects the frame, even if the
                # intensity score is modest. Strong but non-round candidates are still blank-guarded.
                if no_round_droplet_frame and accepted_before_nms:
                    guard_applied = True
                    guard_reason = "no_round_droplet_protector"
                elif no_round_droplet_frame and weak_score_frame:
                    guard_applied = True
                    guard_reason = "weak_score_and_no_round_droplet"
                else:
                    guard_reason = "round_droplet_protected" if round_protector_count >= min_round_droplets else "not_blank_like"
            else:
                # Legacy V5.2 behavior. Kept as an expert fallback.
                if weak_score_frame:
                    guard_applied = True
                    guard_reason = "legacy_weak_score"

            if guard_applied:
                for row in accepted_before_nms:
                    row["accepted"] = False
                    row["rejection_reason"] = "blank_frame_guard"
                    rejected.append(row)
                accepted_before_nms = []

        accepted_sorted = sorted(accepted_before_nms, key=lambda r: float(r["detection_score"]), reverse=True)
        accepted_final: List[Dict[str, object]] = []

        nms_factor = float(params["nms_overlap_factor"])
        for row in accepted_sorted:
            y = float(row["centroid_y_px"])
            x = float(row["centroid_x_px"])
            r = float(row["radius_px"])

            duplicate = False
            for kept in accepted_final:
                ky = float(kept["centroid_y_px"])
                kx = float(kept["centroid_x_px"])
                kr = float(kept["radius_px"])
                if math.hypot(y - ky, x - kx) < nms_factor * max(r, kr):
                    duplicate = True
                    break

            if duplicate:
                row2 = dict(row)
                row2["accepted"] = False
                row2["rejection_reason"] = "duplicate_candidate"
                rejected.append(row2)
            else:
                accepted_final.append(row)

        # Restore frame order roughly by centroid, not by score.
        accepted_final = sorted(accepted_final, key=lambda r: (float(r["centroid_y_px"]), float(r["centroid_x_px"])))
        for i, row in enumerate(accepted_final, start=1):
            row["droplet_number"] = i

        metrics = {
            "candidate_z_threshold": candidate_z_threshold,
            "quality_threshold": q_threshold,
            "min_contrast_required": min_contrast,
            "min_edge_required": min_edge,
            "raw_blob_candidates": int(split_metrics.get("original_blob_candidates", len(blobs))),
            "candidate_blobs_after_splitting": int(split_metrics.get("candidate_blobs_after_splitting", len(blobs))),
            "split_parent_candidates": int(split_metrics.get("split_parent_candidates", 0)),
            "split_child_candidates": int(split_metrics.get("split_child_candidates", 0)),
            "accepted_before_blank_guard_and_nms": int(len(scores)),
            "rejected_candidates": int(len(rejected)),
            "frame_max_detection_score": float(max_score),
            "frame_top5_mean_detection_score": float(top5_mean),
            "frame_max_roundness_score": float(max_roundness),
            "frame_top5_mean_roundness_score": float(top5_roundness_mean),
            "round_droplet_protectors": int(round_protector_count),
            "blank_guard_roundness_threshold_used": float(roundness_threshold),
            "blank_guard_min_round_droplets_used": int(min_round_droplets),
            "blank_frame_guard_applied": bool(guard_applied),
            "blank_frame_guard_reason": guard_reason,
        }

        return accepted_final, rejected, metrics

    def analyze_frame(self, image: np.ndarray, frame_index: int, frame_name: str) -> DetectionOutput:
        params = self.get_params_values()
        preview_gray = display_to_float01(image, gamma=float(params.get("preview_gamma", 0.75)))

        accepted, rejected, metrics = self.detect_candidates(image, frame_index, frame_name, params_override=params)
        overlay = self.make_overlay(
            preview_gray,
            accepted,
            rejected if bool(params["show_rejected_candidates"]) else [],
            bool(params["show_labels"]),
            show_markers=bool(self.preview_markers_visible_var.get()),
        )

        return DetectionOutput(overlay, accepted, rejected, len(accepted), metrics)

    def make_overlay(
        self,
        gray: np.ndarray,
        accepted_rows: List[Dict[str, object]],
        rejected_rows: List[Dict[str, object]],
        show_labels: bool,
        show_markers: bool = True,
    ) -> Image.Image:
        base8 = (np.clip(gray, 0, 1) * 255).astype(np.uint8)
        rgb = np.dstack([base8, base8, base8])
        pil = Image.fromarray(rgb).convert("RGB")

        if not show_markers:
            return pil

        draw = ImageDraw.Draw(pil)

        for row in rejected_rows:
            x = float(row["centroid_x_px"])
            y = float(row["centroid_y_px"])
            r = float(row["radius_px"])
            draw.ellipse((x - r, y - r, x + r, y + r), outline=(255, 165, 0), width=1)

        for row in accepted_rows:
            x = float(row["centroid_x_px"])
            y = float(row["centroid_y_px"])
            r = float(row["radius_px"])
            draw.ellipse((x - r, y - r, x + r, y + r), outline=(255, 0, 0), width=2)
            draw.ellipse((x - 2, y - 2, x + 2, y + 2), outline=(0, 255, 0), width=1)

            if show_labels:
                draw.text((x + r + 2, y + r + 2), str(row["droplet_number"]), fill=(255, 255, 0))

        draw.rectangle((8, 8, 318, 72), fill=(0, 0, 0))
        draw.text((14, 14), "Red circle: accepted droplet", fill=(255, 255, 255))
        draw.text((14, 32), "Green dot: droplet center", fill=(255, 255, 255))
        draw.text((14, 50), f"Count: {len(accepted_rows)}", fill=(255, 255, 255))

        return pil

    def export_excel(self):
        if not self.frames:
            messagebox.showwarning("No images", "Load images before exporting.")
            return

        output = filedialog.asksaveasfilename(
            title="Save Excel analysis",
            defaultextension=".xlsx",
            initialfile="droplet_analysis_v6.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not output:
            return

        try:
            all_rows: List[Dict[str, object]] = []
            summary_rows: List[Dict[str, object]] = []
            original_status = self.status_var.get()

            for i, frame in enumerate(self.frames):
                self.status_var.set(f"Analyzing frame {i + 1}/{len(self.frames)}...")
                self.root.update_idletasks()

                det = self.analyze_frame(frame.image, frame_index=i, frame_name=frame.name)
                all_rows.extend(det.rows)

                summary = {
                    "frame_index": i + 1,
                    "frame_name": frame.name,
                    "temperature_C": extract_temperature_c_from_name(frame.name),
                    "droplet_count": len(det.rows),
                    **det.frame_metrics,
                }

                if det.rows:
                    df_temp = pd.DataFrame(det.rows)
                    summary.update({
                        "mean_area_px2": df_temp["area_px2"].mean(),
                        "median_area_px2": df_temp["area_px2"].median(),
                        "sd_area_px2": df_temp["area_px2"].std(ddof=1) if len(df_temp) > 1 else 0,
                        "mean_diameter_px": df_temp["diameter_px"].mean(),
                        "median_diameter_px": df_temp["diameter_px"].median(),
                        "mean_detection_score": df_temp["detection_score"].mean(),
                        "mean_roundness_score": df_temp["roundness_score"].mean() if "roundness_score" in df_temp.columns else None,
                        "max_roundness_score": df_temp["roundness_score"].max() if "roundness_score" in df_temp.columns else None,
                        "mean_local_contrast": df_temp["local_contrast"].mean(),
                        "mean_edge_sharpness": df_temp["edge_sharpness"].mean(),
                    })

                    if "area_um2" in df_temp.columns and df_temp["area_um2"].notna().any():
                        summary.update({
                            "mean_area_um2": df_temp["area_um2"].mean(),
                            "median_area_um2": df_temp["area_um2"].median(),
                            "sd_area_um2": df_temp["area_um2"].std(ddof=1) if len(df_temp) > 1 else 0,
                            "mean_diameter_um": df_temp["diameter_um"].mean(),
                            "median_diameter_um": df_temp["diameter_um"].median(),
                        })
                else:
                    summary.update({
                        "mean_area_px2": None,
                        "median_area_px2": None,
                        "sd_area_px2": None,
                        "mean_diameter_px": None,
                        "median_diameter_px": None,
                        "mean_detection_score": None,
                        "mean_roundness_score": None,
                        "max_roundness_score": None,
                        "mean_local_contrast": None,
                        "mean_edge_sharpness": None,
                    })

                summary_rows.append(summary)

            if all_rows:
                droplet_df = pd.DataFrame(all_rows)
            else:
                droplet_df = pd.DataFrame(columns=[
                    "frame_index", "frame_name", "temperature_C", "droplet_number",
                    "centroid_x_px", "centroid_y_px", "radius_px", "diameter_px",
                    "area_px2", "perimeter_px", "radius_um", "diameter_um",
                    "area_um2", "perimeter_um", "local_contrast",
                    "edge_sharpness", "signal_z", "detection_score",
                    "roundness_score", "shape_circularity", "shape_solidity",
                    "shape_eccentricity", "angular_coverage", "angular_uniformity",
                ])

            summary_df = pd.DataFrame(summary_rows)

            # Keep temperature_C immediately after frame_name even if pandas changes column order.
            for df in (droplet_df, summary_df):
                if "frame_name" in df.columns and "temperature_C" in df.columns:
                    cols = list(df.columns)
                    cols.remove("temperature_C")
                    insert_at = cols.index("frame_name") + 1
                    cols.insert(insert_at, "temperature_C")
                    df = df[cols]
                    if "droplet_number" in df.columns:
                        droplet_df = df
                    else:
                        summary_df = df

            params = self.get_params_values()
            params_rows = [{"parameter": key, "value": value} for key, value in params.items()]
            params_rows.extend([
                {"parameter": "algorithm_note", "value": "V6.1 detects dark brightfield droplets by local background subtraction, multiscale blob detection, optional close-droplet splitting, intensity gates, shape-first filters, duplicate removal, and roundness-protected blank-frame guarding."},
                {"parameter": "performance_note", "value": "V6.1 caches preprocessed signal images and detection results for preview use. Marker display, labels, rejected-candidate display, and preview gamma do not re-run detection."},
                {"parameter": "linked_controls_note", "value": "In linked mode, Droplet sensitivity, Shape strictness, and Close-droplet separation drive the effective Advanced thresholds for candidate detection, contrast, edge sharpness, roundness, circularity, angular coverage, splitting, duplicate removal, and blank guarding."},
                {"parameter": "shape_filter_note", "value": "Shape-first filters reject candidates with low roundness, low circularity, high eccentricity, incomplete angular edge coverage, or poor angular uniformity."},
                {"parameter": "split_note", "value": "Close-droplet splitting uses local peak detection inside large blob candidates before duplicate removal. Splitting diagnostics are recorded in per_frame_summary."},
                {"parameter": "preview_marker_note", "value": "The Hide/Show markers toolbar button affects only the preview overlay. It does not change droplet detection, measurements, or Excel export values."},
                {"parameter": "temperature_extraction_note", "value": "temperature_C is extracted from filename tokens like 25_76C_2W-oT_cropped, where 76C is exported as 76. If no C/°C token is found, the value is left blank."},
                {"parameter": "blank_guard_roundness_note", "value": "Blank-frame guard is protected by roundness. If enough candidates have roundness_score above blank_guard_roundness_threshold, the frame is not blank-guarded. Non-round debris does not protect the frame."},
                {"parameter": "area_note", "value": "Area is calculated from the detected circular droplet radius, not from an irregular threshold mask."},
                {"parameter": "red_overlay", "value": "Accepted droplet circle."},
                {"parameter": "green_overlay", "value": "Detected droplet center."},
            ])
            params_df = pd.DataFrame(params_rows)

            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                droplet_df.to_excel(writer, index=False, sheet_name="per_droplet")
                summary_df.to_excel(writer, index=False, sheet_name="per_frame_summary")
                params_df.to_excel(writer, index=False, sheet_name="parameters")

                for sheet_name in writer.book.sheetnames:
                    ws = writer.book[sheet_name]
                    for col in ws.columns:
                        max_len = 0
                        col_letter = col[0].column_letter
                        for cell in col:
                            val = "" if cell.value is None else str(cell.value)
                            max_len = max(max_len, len(val))
                        ws.column_dimensions[col_letter].width = min(max(10, max_len + 2), 46)

            self.status_var.set(original_status)
            messagebox.showinfo("Export complete", f"Saved Excel file:\n{output}")

        except Exception as exc:
            tb = traceback.format_exc()
            self.status_var.set(f"Export failed: {exc}")
            messagebox.showerror("Export failed", f"{exc}\n\nDetails:\n{tb}")


def main():
    root = tk.Tk()

    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass

    DropletCounterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
