import os
import re
import sys
import subprocess
import traceback
import warnings
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# -------------------------
# Explicit dependency check (no installation at startup)
# -------------------------
REQUIRED_PACKAGES = [
    "pandas",
    "numpy",
    "matplotlib",
    "scipy",
    "openpyxl",
    "xlsxwriter",
]


def ensure_packages():
    missing = []
    for package in REQUIRED_PACKAGES:
        try:
            __import__(package)
        except ImportError:
            missing.append(package)
    if missing:
        raise RuntimeError(
            "Missing dependencies: " + ", ".join(missing)
            + ". Install the repository requirements.txt with this Python interpreter "
            "before starting the program. No packages were installed automatically."
        )


ensure_packages()

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.optimize import curve_fit

warnings.filterwarnings("ignore", category=RuntimeWarning)


BASE_DIR = Path(__file__).resolve().parent
PLOTS_DIR = BASE_DIR / "plots"
PLOTS_DIR.mkdir(exist_ok=True)

SUPPORTED_SUFFIXES = {".csv"}  # Deliberately CSV only. Other files are ignored.
CHANNEL_PATTERN = re.compile(r"(Channel\d+)_R(\d+)_(.+)")
FAST_HALF_LIFE_LIMIT_S = 20.0
FAST_RECOVERY_FIT_WINDOW_S = 90.0
MIN_DENOMINATOR_ABS = 1e-12


# ============================================================
# Data containers
# ============================================================
@dataclass
class SampleInput:
    name: str
    folder: Path
    files: List[Path]
    output_excel: Path

    @property
    def total_csv_count(self) -> int:
        return len(self.files)


@dataclass
class SampleRunConfig:
    sample: SampleInput
    fit_mode: str
    fixed_fit_window_s: Optional[float]


class FRAPExperiment:
    def __init__(self, file_name: str):
        self.file_name = file_name
        self.time_seconds: Optional[np.ndarray] = None
        self.bleach_time_seconds: Optional[float] = None
        self.channels: Dict[str, Dict[str, np.ndarray]] = {}
        self.processed: Dict[str, np.ndarray] = {}
        self.normalized: Dict[str, pd.DataFrame] = {}
        self.fit_results: Dict[str, Dict[str, float]] = {}


@dataclass
class FRAPFrameTimingData:
    path: Path
    source_unit: str
    frame_numbers: np.ndarray
    raw_time_seconds: np.ndarray
    time_from_start_seconds: np.ndarray
    markers: List[str]
    bleach_index: Optional[int]


@dataclass
class FRAPFrameMatch:
    requested_time_s: float
    frame_number: int
    exact_time_from_reference_s: float
    time_from_measurement_start_s: float
    raw_csv_time_s: float
    difference_s: float
    relation: str


# ============================================================
# File/sample detection
# ============================================================
def safe_filename(text: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_.\- ]+", "_", text).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned or "FRAP_sample"


def list_csv_files(folder: Path) -> List[Path]:
    return sorted([p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in SUPPORTED_SUFFIXES])


def detect_samples(chosen_folder: Path) -> List[SampleInput]:
    """
    Detects FRAP samples in two layouts:
    1. The chosen folder directly contains CSV files: one sample named after the chosen folder.
    2. The chosen folder contains subfolders: each subfolder with CSV files is one sample.

    If both direct CSV files and sample subfolders exist, both are analyzed.
    Non-CSV files are ignored.
    """
    chosen_folder = Path(chosen_folder).resolve()
    if not chosen_folder.exists() or not chosen_folder.is_dir():
        raise FileNotFoundError(f"Selected folder does not exist: {chosen_folder}")

    samples: List[SampleInput] = []

    direct_files = list_csv_files(chosen_folder)
    if direct_files:
        sample_name = chosen_folder.name
        samples.append(
            SampleInput(
                name=sample_name,
                folder=chosen_folder,
                files=direct_files,
                output_excel=chosen_folder / f"{safe_filename(sample_name)}_summary.xlsx",
            )
        )

    for child in sorted([p for p in chosen_folder.iterdir() if p.is_dir()]):
        child_files = list_csv_files(child)
        if not child_files:
            continue
        sample_name = child.name
        samples.append(
            SampleInput(
                name=sample_name,
                folder=child,
                files=child_files,
                output_excel=chosen_folder / f"{safe_filename(sample_name)}_summary.xlsx",
            )
        )

    return samples


# ============================================================
# Analysis core
# ============================================================
def unit_to_seconds_multiplier(unit_value: str) -> float:
    if unit_value is None:
        return 1.0
    text = str(unit_value).strip().lower()
    mapping = {
        "ms": 1e-3,
        "msec": 1e-3,
        "millisecond": 1e-3,
        "milliseconds": 1e-3,
        "s": 1.0,
        "sec": 1.0,
        "secs": 1.0,
        "second": 1.0,
        "seconds": 1.0,
        "min": 60.0,
        "mins": 60.0,
        "minute": 60.0,
        "minutes": 60.0,
        "h": 3600.0,
        "hr": 3600.0,
        "hrs": 3600.0,
        "hour": 3600.0,
        "hours": 3600.0,
    }
    return mapping.get(text, 1.0)


def read_input_file(path: Path) -> pd.DataFrame:
    if path.suffix.lower() != ".csv":
        raise ValueError(f"Only CSV files are supported. Ignored file: {path.name}")
    return pd.read_csv(path, header=0)


def pick_relevant_intensity_columns(columns: List[str]) -> List[str]:
    selected = []
    for col in columns:
        left = str(col).split("::")[0]
        if not left.startswith("Channel"):
            continue
        if "_IntensityMean" in left and "IntensityMeanThrs" not in left:
            selected.append(col)
    return selected


def parse_channel_and_roi(column_name: str) -> Optional[Tuple[str, str]]:
    left = str(column_name).split("::")[0]
    match = CHANNEL_PATTERN.match(left)
    if not match:
        return None
    channel = match.group(1)
    roi = f"R{match.group(2)}"
    return channel, roi


def asymptotic_exp(x: np.ndarray, a: float, b: float, k: float) -> np.ndarray:
    return a - b * np.exp(k * x)


def parse_frap_file(path: Path) -> FRAPExperiment:
    raw = read_input_file(path)
    exp = FRAPExperiment(path.stem)

    if raw.shape[0] < 2 or raw.shape[1] < 2:
        raise ValueError(f"File has an unexpected format: {path.name}")

    unit_text = raw.iloc[0, 0]
    multiplier = unit_to_seconds_multiplier(unit_text)

    data = raw.iloc[1:, :].copy().reset_index(drop=True)
    data.columns = [str(c).strip() for c in data.columns]

    time_col = data.columns[0]
    marker_col = data.columns[1]

    data[time_col] = pd.to_numeric(data[time_col], errors="coerce")
    data = data.dropna(subset=[time_col]).reset_index(drop=True)
    data["time_seconds"] = data[time_col].astype(float) * multiplier

    bleach_mask = data[marker_col].astype(str).str.contains("bleach", case=False, na=False)
    if not bleach_mask.any():
        raise ValueError(f"Could not find bleaching row in file: {path.name}")

    bleach_idx = int(np.flatnonzero(bleach_mask.to_numpy())[0])
    bleach_time = float(data.loc[bleach_idx, "time_seconds"])

    exp.time_seconds = data["time_seconds"].to_numpy(dtype=float)
    exp.bleach_time_seconds = bleach_time

    intensity_cols = pick_relevant_intensity_columns(list(data.columns))
    if not intensity_cols:
        raise ValueError(f"No usable intensity columns found in file: {path.name}")

    for col in intensity_cols:
        parsed = parse_channel_and_roi(col)
        if parsed is None:
            continue
        channel, roi = parsed
        exp.channels.setdefault(channel, {})[roi] = pd.to_numeric(data[col], errors="coerce").to_numpy(dtype=float)

    return exp


def parse_requested_timepoints(text: str) -> List[float]:
    """Parse comma, semicolon, whitespace, or newline separated times in seconds."""
    tokens = [token for token in re.split(r"[,;\s]+", str(text).strip()) if token]
    if not tokens:
        raise ValueError("Enter at least one requested timepoint.")

    values: List[float] = []
    for token in tokens:
        try:
            value = float(token.replace(",", "."))
        except ValueError as exc:
            raise ValueError(f"Invalid timepoint: {token}") from exc
        if not np.isfinite(value) or value < 0:
            raise ValueError("Requested timepoints must be finite numbers greater than or equal to 0.")
        values.append(value)
    return values


def format_time_value(value: float) -> str:
    if abs(value) < 5e-10:
        value = 0.0
    return f"{value:.6f}".rstrip("0").rstrip(".")


def parse_frap_frame_timing(path: Path) -> FRAPFrameTimingData:
    """
    Parse the native FRAP CSV structure used by the main analysis pipeline.

    Expected structure:
    - CSV header on row 1
    - Unit in the first cell of row 2, for example "min"
    - Acquired frames from row 3 onward
    - Relative acquisition time in the first column
    - Bleaching marker text in the second column
    """
    raw = read_input_file(Path(path))
    if raw.shape[0] < 2 or raw.shape[1] < 2:
        raise ValueError(f"File has an unexpected FRAP format: {Path(path).name}")

    source_unit = str(raw.iloc[0, 0]).strip()
    multiplier = unit_to_seconds_multiplier(source_unit)

    data = raw.iloc[1:, :].copy().reset_index(drop=True)
    time_col = data.columns[0]
    marker_col = data.columns[1]

    numeric_time = pd.to_numeric(data[time_col], errors="coerce")
    valid_mask = numeric_time.notna().to_numpy()
    if not np.any(valid_mask):
        raise ValueError("No numeric acquisition times were found after the unit row.")

    source_frame_numbers = np.arange(1, len(data) + 1, dtype=int)
    frame_numbers = source_frame_numbers[valid_mask]
    raw_time_seconds = numeric_time.to_numpy(dtype=float)[valid_mask] * multiplier
    markers = data[marker_col].fillna("").astype(str).to_numpy()[valid_mask].tolist()

    if len(raw_time_seconds) == 0:
        raise ValueError("No valid acquisition frames were found.")

    time_from_start = raw_time_seconds - raw_time_seconds[0]
    bleach_candidates = [
        index for index, marker in enumerate(markers)
        if "bleach" in marker.casefold()
    ]
    bleach_index = bleach_candidates[0] if bleach_candidates else None

    return FRAPFrameTimingData(
        path=Path(path),
        source_unit=source_unit,
        frame_numbers=frame_numbers,
        raw_time_seconds=raw_time_seconds,
        time_from_start_seconds=time_from_start,
        markers=markers,
        bleach_index=bleach_index,
    )


def match_frap_timepoints(
    timing: FRAPFrameTimingData,
    requested_times_s: List[float],
    reference_mode: str,
    first_frame_number: int = 1,
) -> Tuple[List[FRAPFrameMatch], str]:
    if reference_mode == "bleach":
        if timing.bleach_index is None:
            raise ValueError("No bleaching marker was found in the marker column.")
        reference_index = timing.bleach_index
        reference_label = (
            f"bleach frame {first_frame_number + int(timing.frame_numbers[reference_index]) - 1}, "
            f"{format_time_value(timing.time_from_start_seconds[reference_index])} s from measurement start"
        )
    else:
        reference_index = 0
        reference_label = (
            f"measurement start, frame {first_frame_number + int(timing.frame_numbers[0]) - 1}"
        )

    reference_raw_time = float(timing.raw_time_seconds[reference_index])
    relative_times = timing.raw_time_seconds - reference_raw_time
    results: List[FRAPFrameMatch] = []

    for requested in requested_times_s:
        differences = np.abs(relative_times - requested)
        selected_index = int(np.argmin(differences))
        exact_relative = float(relative_times[selected_index])
        difference = exact_relative - requested

        tolerance = 1e-9
        if requested > float(relative_times[-1]) + tolerance:
            relation = "After available range"
        elif abs(difference) <= tolerance:
            relation = "Exact"
        elif difference < 0:
            relation = "Before"
        else:
            relation = "After"

        results.append(
            FRAPFrameMatch(
                requested_time_s=float(requested),
                frame_number=first_frame_number + int(timing.frame_numbers[selected_index]) - 1,
                exact_time_from_reference_s=exact_relative,
                time_from_measurement_start_s=float(timing.time_from_start_seconds[selected_index]),
                raw_csv_time_s=float(timing.raw_time_seconds[selected_index]),
                difference_s=difference,
                relation=relation,
            )
        )

    return results, reference_label


def compute_processed_trace(exp: FRAPExperiment, channel: str) -> np.ndarray:
    traces = exp.channels[channel]
    time = exp.time_seconds
    bleach_time = exp.bleach_time_seconds

    r1 = traces.get("R1")
    r2 = traces.get("R2")
    r3 = traces.get("R3")

    if r1 is None:
        raise ValueError(f"{exp.file_name} / {channel}: missing R1")

    pre_mask = time < bleach_time
    if not np.any(pre_mask):
        raise ValueError(f"{exp.file_name} / {channel}: no pre-bleach points detected")

    r2_pre = np.nanmean(r2[pre_mask]) if r2 is not None else np.nan
    r3_pre = np.nanmean(r3[pre_mask]) if r3 is not None else 0.0

    if r2 is not None and r3 is not None:
        numerator = (r1 - r3) * (r2_pre - r3_pre)
        denominator = r2 - r3
        processed = np.divide(
            numerator,
            denominator,
            out=np.full_like(numerator, np.nan, dtype=float),
            where=np.abs(denominator) > MIN_DENOMINATOR_ABS,
        )
    elif r2 is not None and r3 is None:
        processed = np.divide(
            r1 * r2_pre,
            r2,
            out=np.full_like(r1, np.nan, dtype=float),
            where=np.abs(r2) > MIN_DENOMINATOR_ABS,
        )
    elif r2 is None and r3 is not None:
        processed = r1 - r3
    else:
        processed = r1.copy()

    return processed


def normalize_trace(time_seconds: np.ndarray, bleach_time_seconds: float, processed: np.ndarray) -> pd.DataFrame:
    time_from_bleach = time_seconds - bleach_time_seconds
    post_mask = time_from_bleach >= 0

    x = time_from_bleach[post_mask]
    y = processed[post_mask]

    pre_mask = time_seconds < bleach_time_seconds
    pre_bleach_max = np.nanmax(processed[pre_mask]) if np.any(pre_mask) else np.nan
    post_bleach_min = np.nanmin(y) if len(y) > 0 else np.nan

    denominator = pre_bleach_max - post_bleach_min
    if np.isnan(denominator) or abs(denominator) < MIN_DENOMINATOR_ABS:
        normalized = np.full_like(y, np.nan, dtype=float)
    else:
        normalized = 100.0 * (y - post_bleach_min) / denominator

    return pd.DataFrame(
        {
            "time_from_bleach_s": x,
            "normalized_intensity": normalized,
        }
    )


def _fit_half_life_once(x: np.ndarray, y: np.ndarray, fit_window_s: Optional[float] = None) -> Dict[str, float]:
    valid = np.isfinite(x) & np.isfinite(y) & (x >= 0)
    x = x[valid]
    y = y[valid]

    if fit_window_s is not None:
        within_window = x <= fit_window_s
        x = x[within_window]
        y = y[within_window]

    if len(x) < 4:
        return {
            "a": np.nan,
            "b": np.nan,
            "k": np.nan,
            "c": np.nan,
            "half_life_s": np.nan,
            "fit_r2": np.nan,
            "fit_timeframe_s": float(np.nanmax(x)) if len(x) > 0 else np.nan,
        }

    a0 = max(80.0, float(np.nanmax(y)))
    b0 = max(1.0, a0 - float(np.nanmin(y)))
    total_time = max(float(np.nanmax(x) - np.nanmin(x)), 1.0)
    k0 = -1.0 / total_time

    try:
        popt, _ = curve_fit(
            asymptotic_exp,
            x,
            y,
            p0=[a0, b0, k0],
            bounds=([0.0, 0.0, -np.inf], [200.0, 300.0, -1e-9]),
            maxfev=50000,
        )
        a, b, k = popt
        y_fit = asymptotic_exp(x, a, b, k)
        ss_res = float(np.nansum((y - y_fit) ** 2))
        ss_tot = float(np.nansum((y - np.nanmean(y)) ** 2))
        r2 = np.nan if ss_tot == 0 else 1.0 - ss_res / ss_tot

        c = float(np.exp(k))
        ratio = (0.5 * a) / b if b != 0 else np.nan
        if not np.isfinite(ratio) or ratio <= 0 or k == 0:
            half_life = np.nan
        else:
            half_life = float(np.log(ratio) / k)
            if half_life < 0:
                half_life = np.nan

        used_timeframe = float(np.nanmax(x))
        return {
            "a": float(a),
            "b": float(b),
            "k": float(k),
            "c": c,
            "half_life_s": half_life,
            "fit_r2": float(r2),
            "fit_timeframe_s": used_timeframe,
        }
    except Exception:
        return {
            "a": np.nan,
            "b": np.nan,
            "k": np.nan,
            "c": np.nan,
            "half_life_s": np.nan,
            "fit_r2": np.nan,
            "fit_timeframe_s": float(np.nanmax(x)) if len(x) > 0 else np.nan,
        }


def fit_half_life(norm_df: pd.DataFrame, fit_mode: str, fixed_fit_window_s: Optional[float]) -> Dict[str, float]:
    x = norm_df["time_from_bleach_s"].to_numpy(dtype=float)
    y = norm_df["normalized_intensity"].to_numpy(dtype=float)

    if fit_mode == "fixed" and fixed_fit_window_s is not None:
        result = _fit_half_life_once(x, y, fit_window_s=fixed_fit_window_s)
        result["refit_due_to_fast_recovery"] = False
        return result

    first_fit = _fit_half_life_once(x, y, fit_window_s=None)
    half_life = first_fit.get("half_life_s", np.nan)

    if np.isfinite(half_life) and 0 <= half_life <= FAST_HALF_LIFE_LIMIT_S:
        refit = _fit_half_life_once(x, y, fit_window_s=FAST_RECOVERY_FIT_WINDOW_S)
        refit["refit_due_to_fast_recovery"] = True
        return refit

    first_fit["refit_due_to_fast_recovery"] = False
    return first_fit


# ============================================================
# Sheet/plot builders
# ============================================================
def build_raw_sheet(experiments: List[FRAPExperiment], selected_channels: List[str]) -> pd.DataFrame:
    """
    Builds the raw-data sheet using independent Series objects for each file/time vector.
    This is deliberate: files with shorter or longer acquisition lengths keep their own X-axis
    and do not inherit the X-axis length from whichever file happens to appear first.
    """
    columns: Dict[str, pd.Series] = {}
    roi_order = ["R1", "R2", "R3"]

    for channel in selected_channels:
        for roi in roi_order:
            for exp in experiments:
                if channel in exp.channels and roi in exp.channels[channel]:
                    columns[f"time_s_{channel}_{roi}_{exp.file_name}"] = pd.Series(exp.time_seconds)
                    columns[f"{channel}_{roi}_{exp.file_name}"] = pd.Series(exp.channels[channel][roi])

    return pd.DataFrame(columns)


def build_normalized_sheet(experiments: List[FRAPExperiment], selected_channels: List[str]) -> pd.DataFrame:
    """
    Builds the normalized-data sheet using each file's own post-bleach time vector.
    No file is used as the master X-axis.
    """
    columns: Dict[str, pd.Series] = {}

    for channel in selected_channels:
        for exp in experiments:
            if channel not in exp.normalized:
                continue
            norm_df = exp.normalized[channel].copy().reset_index(drop=True)
            columns[f"time_from_bleach_s_{channel}_{exp.file_name}"] = pd.Series(norm_df["time_from_bleach_s"])
            columns[f"data_{channel}_{exp.file_name}"] = pd.Series(norm_df["normalized_intensity"])

    return pd.DataFrame(columns)


def build_normalized_mean_sd_sheet(experiments: List[FRAPExperiment], selected_channels: List[str]) -> pd.DataFrame:
    """
    Builds the averaged normalized recovery table shown next to the normalized-data plot.
    The average and SD are calculated on the same union/interpolated time axis used for the plot.
    SD is left blank when only one finite trace contributes to a time point.
    """
    columns: Dict[str, pd.Series] = {}

    for channel in selected_channels:
        union_x, matrix = interpolate_to_union_time(experiments, channel)
        if union_x.size == 0 or matrix.size == 0:
            continue

        mean_y = np.nanmean(matrix, axis=0)
        counts = np.sum(np.isfinite(matrix), axis=0)
        std_y = np.full_like(mean_y, np.nan, dtype=float)
        valid_for_sd = counts > 1
        if np.any(valid_for_sd):
            std_y[valid_for_sd] = np.nanstd(matrix[:, valid_for_sd], axis=0, ddof=1)

        columns[f"time_from_bleach_s_{channel}_mean"] = pd.Series(union_x)
        columns[f"average_normalized_intensity_{channel}"] = pd.Series(mean_y)
        columns[f"SD_{channel}"] = pd.Series(std_y)

    return pd.DataFrame(columns)


def build_half_life_sheet(experiments: List[FRAPExperiment], selected_channels: List[str]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    rows = []

    for exp in experiments:
        row = {"file_name": exp.file_name}
        for channel in selected_channels:
            result = exp.fit_results.get(channel, {})
            row[f"{channel}_half_life_s"] = result.get("half_life_s", np.nan)
            row[f"{channel}_fit_timeframe_s"] = result.get("fit_timeframe_s", np.nan)
            row[f"{channel}_R2"] = result.get("fit_r2", np.nan)
            row[f"{channel}_refit_fast_recovery"] = result.get("refit_due_to_fast_recovery", False)
        rows.append(row)

    detailed = pd.DataFrame(rows)

    summary_rows = []
    for channel in selected_channels:
        half_life_col = f"{channel}_half_life_s"
        timeframe_col = f"{channel}_fit_timeframe_s"

        values = detailed[half_life_col].to_numpy(dtype=float) if half_life_col in detailed.columns else np.array([], dtype=float)
        windows = detailed[timeframe_col].to_numpy(dtype=float) if timeframe_col in detailed.columns else np.array([], dtype=float)

        finite_values = np.isfinite(values)
        finite_windows = np.isfinite(windows)

        summary_rows.append(
            {
                "channel": channel,
                "mean_half_life_s": np.nanmean(values) if np.any(finite_values) else np.nan,
                "std_half_life_s": np.nanstd(values, ddof=1) if np.sum(finite_values) > 1 else np.nan,
                "mean_fit_timeframe_s": np.nanmean(windows) if np.any(finite_windows) else np.nan,
                "n": int(np.sum(finite_values)),
            }
        )

    summary = pd.DataFrame(summary_rows, columns=["channel", "mean_half_life_s", "std_half_life_s", "mean_fit_timeframe_s", "n"])
    return detailed, summary


def create_raw_r1_plot(experiments: List[FRAPExperiment], channel: str, out_path: Path):
    plt.figure(figsize=(8, 5))
    plotted = False
    for exp in experiments:
        if channel in exp.channels and "R1" in exp.channels[channel]:
            plt.plot(exp.time_seconds, exp.channels[channel]["R1"], linewidth=1.5, label=exp.file_name)
            plotted = True
    if not plotted:
        plt.close()
        return
    plt.xlabel("Time (s)")
    plt.ylabel("Raw R1 intensity")
    plt.title(f"Raw R1 traces - {channel}")
    plt.legend(fontsize=8)
    plt.tight_layout()
    plt.savefig(out_path, dpi=300)
    plt.close()


def interpolate_to_union_time(experiments: List[FRAPExperiment], channel: str) -> Tuple[np.ndarray, np.ndarray]:
    all_x = []
    curves = []

    for exp in experiments:
        if channel not in exp.normalized:
            continue
        df = exp.normalized[channel]
        x = df["time_from_bleach_s"].to_numpy(dtype=float)
        y = df["normalized_intensity"].to_numpy(dtype=float)
        valid = np.isfinite(x) & np.isfinite(y)
        x = x[valid]
        y = y[valid]
        if len(x) < 2:
            continue
        order = np.argsort(x)
        x = x[order]
        y = y[order]
        all_x.extend(x.tolist())
        curves.append((x, y))

    if not all_x or not curves:
        return np.array([]), np.empty((0, 0))

    union_x = np.array(sorted(set(float(v) for v in all_x)), dtype=float)
    matrix = np.full((len(curves), len(union_x)), np.nan, dtype=float)

    for i, (x, y) in enumerate(curves):
        matrix[i, :] = np.interp(union_x, x, y, left=np.nan, right=np.nan)
        matrix[i, union_x < np.nanmin(x)] = np.nan
        matrix[i, union_x > np.nanmax(x)] = np.nan

    return union_x, matrix


def create_normalized_mean_plot(experiments: List[FRAPExperiment], channel: str, out_path: Path):
    union_x, matrix = interpolate_to_union_time(experiments, channel)
    if union_x.size == 0 or matrix.size == 0:
        return

    mean_y = np.nanmean(matrix, axis=0)
    counts = np.sum(np.isfinite(matrix), axis=0)
    std_y = np.full_like(mean_y, np.nan, dtype=float)
    valid_for_sd = counts > 1
    if np.any(valid_for_sd):
        std_y[valid_for_sd] = np.nanstd(matrix[:, valid_for_sd], axis=0, ddof=1)

    plt.figure(figsize=(8, 5))
    if np.any(np.isfinite(std_y)):
        plt.fill_between(union_x, mean_y - std_y, mean_y + std_y, alpha=0.2)
    plt.plot(union_x, mean_y, linewidth=2.5)
    plt.xlabel("Time from bleach (s)")
    plt.ylabel("Normalized intensity")
    plt.title(f"Normalized FRAP recovery - {channel} (mean ± SD)")
    plt.tight_layout()
    plt.savefig(out_path, dpi=300)
    plt.close()


def write_output_excel(
    output_excel: Path,
    raw_df: pd.DataFrame,
    norm_df: pd.DataFrame,
    norm_mean_sd_df: pd.DataFrame,
    half_life_df: pd.DataFrame,
    summary_df: pd.DataFrame,
    raw_plot_paths: Dict[str, Path],
    norm_plot_paths: Dict[str, Path],
):
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        raw_df.to_excel(writer, sheet_name="raw_data", index=False)
        norm_df.to_excel(writer, sheet_name="normalized_data", index=False)
        mean_sd_start_col = norm_df.shape[1] + 2
        if not norm_mean_sd_df.empty:
            norm_mean_sd_df.to_excel(
                writer,
                sheet_name="normalized_data",
                index=False,
                startrow=0,
                startcol=mean_sd_start_col,
            )
        half_life_df.to_excel(writer, sheet_name="half_life", index=False, startrow=0)
        summary_df.to_excel(writer, sheet_name="half_life", index=False, startrow=len(half_life_df) + 3)

        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True})

        for sheet_name in ["raw_data", "normalized_data", "half_life"]:
            worksheet = writer.sheets[sheet_name]
            worksheet.set_row(0, None, header_fmt)

        raw_ws = writer.sheets["raw_data"]
        norm_ws = writer.sheets["normalized_data"]
        hl_ws = writer.sheets["half_life"]

        raw_ws.set_column(0, raw_df.shape[1], 24)
        norm_ws.set_column(0, norm_df.shape[1], 24)
        if not norm_mean_sd_df.empty:
            norm_ws.set_column(mean_sd_start_col, mean_sd_start_col + norm_mean_sd_df.shape[1], 28)
        hl_ws.set_column(0, max(half_life_df.shape[1], summary_df.shape[1]), 22)

        start_row = 2
        plot_col = raw_df.shape[1] + 2
        for channel, plot_path in raw_plot_paths.items():
            if plot_path.exists():
                raw_ws.insert_image(start_row, plot_col, str(plot_path), {"x_scale": 0.9, "y_scale": 0.9})
                start_row += 24

        start_row = 2
        plot_col = norm_df.shape[1] + norm_mean_sd_df.shape[1] + 4
        for channel, plot_path in norm_plot_paths.items():
            if plot_path.exists():
                norm_ws.insert_image(start_row, plot_col, str(plot_path), {"x_scale": 0.9, "y_scale": 0.9})
                start_row += 24

        hl_ws.write(len(half_life_df) + 2, 0, "Channel summary")


def detect_channels_in_files(files: List[Path]) -> List[str]:
    detected = set()
    for path in files:
        raw = read_input_file(path)
        headers = [str(c).strip() for c in raw.columns]
        for header in headers:
            parsed = parse_channel_and_roi(header)
            if parsed is not None:
                detected.add(parsed[0])

    def channel_sort_key(channel: str):
        numbers = re.findall(r"\d+", channel)
        return int(numbers[0]) if numbers else 999999

    return sorted(detected, key=channel_sort_key)


def run_analysis_for_sample(
    sample: SampleInput,
    channel_selection: str,
    fit_mode: str,
    fixed_fit_window_s: Optional[float],
) -> str:
    files = sorted(sample.files)
    if not files:
        raise FileNotFoundError(f"No CSV input files found in sample: {sample.name}")

    detected_channels = detect_channels_in_files(files)
    if not detected_channels:
        raise ValueError(f"No channels were detected in sample: {sample.name}")

    if channel_selection == "auto":
        selected_channels = detected_channels
    else:
        requested_count = int(channel_selection)
        selected_channels = [f"Channel{i}" for i in range(1, requested_count + 1) if f"Channel{i}" in detected_channels]

    if not selected_channels:
        raise ValueError(f"The selected channel setting does not match any channels found in sample: {sample.name}")

    experiments: List[FRAPExperiment] = []
    for path in files:
        exp = parse_frap_file(path)
        for channel in selected_channels:
            if channel not in exp.channels:
                continue
            processed = compute_processed_trace(exp, channel)
            exp.processed[channel] = processed
            norm_df = normalize_trace(exp.time_seconds, exp.bleach_time_seconds, processed)
            exp.normalized[channel] = norm_df
            exp.fit_results[channel] = fit_half_life(norm_df, fit_mode=fit_mode, fixed_fit_window_s=fixed_fit_window_s)
        experiments.append(exp)

    raw_df = build_raw_sheet(experiments, selected_channels)
    norm_df = build_normalized_sheet(experiments, selected_channels)
    norm_mean_sd_df = build_normalized_mean_sd_sheet(experiments, selected_channels)
    half_life_df, summary_df = build_half_life_sheet(experiments, selected_channels)

    raw_plot_paths: Dict[str, Path] = {}
    norm_plot_paths: Dict[str, Path] = {}
    sample_plot_prefix = safe_filename(sample.name)
    for channel in selected_channels:
        raw_plot = PLOTS_DIR / f"{sample_plot_prefix}_raw_R1_{channel}.png"
        norm_plot = PLOTS_DIR / f"{sample_plot_prefix}_normalized_{channel}.png"
        create_raw_r1_plot(experiments, channel, raw_plot)
        create_normalized_mean_plot(experiments, channel, norm_plot)
        raw_plot_paths[channel] = raw_plot
        norm_plot_paths[channel] = norm_plot

    write_output_excel(
        output_excel=sample.output_excel,
        raw_df=raw_df,
        norm_df=norm_df,
        norm_mean_sd_df=norm_mean_sd_df,
        half_life_df=half_life_df,
        summary_df=summary_df,
        raw_plot_paths=raw_plot_paths,
        norm_plot_paths=norm_plot_paths,
    )

    fit_description = "automatic" if fit_mode == "auto" else f"fixed ({fixed_fit_window_s:.2f} s)"
    return (
        f"Sample: {sample.name}\n"
        f"CSV files analyzed: {len(files)}\n"
        f"Channels used: {', '.join(selected_channels)}\n"
        f"Fit mode: {fit_description}\n"
        f"Output: {sample.output_excel}"
    )


def open_file_with_default_app(path: Path):
    path = Path(path)
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception as exc:
        raise RuntimeError(f"Could not open output file: {path}\n{exc}")


# ============================================================
# GUI
# ============================================================
class FileSelectionDialog(tk.Toplevel):
    def __init__(self, parent, sample: SampleInput, initially_selected: List[Path]):
        super().__init__(parent)
        self.title(f"Choose CSV files - {sample.name}")
        self.geometry("560x420")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()

        self.sample = sample
        selected_set = {Path(p).resolve() for p in initially_selected}
        self.file_vars: Dict[Path, tk.BooleanVar] = {
            path: tk.BooleanVar(value=path.resolve() in selected_set)
            for path in sample.files
        }
        self.result: Optional[List[Path]] = None

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._cancel)

    def _build_ui(self):
        outer = ttk.Frame(self, padding=12)
        outer.pack(fill="both", expand=True)

        ttk.Label(
            outer,
            text=(
                f"Sample: {self.sample.name}\n"
                "Checked CSV files will be included in the analysis. "
                "By default, all CSV files are selected."
            ),
            justify="left",
        ).pack(anchor="w", pady=(0, 10))

        button_row_top = ttk.Frame(outer)
        button_row_top.pack(fill="x", pady=(0, 8))
        ttk.Button(button_row_top, text="Select all", command=self._select_all).pack(side="left")
        ttk.Button(button_row_top, text="Select none", command=self._select_none).pack(side="left", padx=(8, 0))

        list_frame = ttk.LabelFrame(outer, text="CSV files", padding=8)
        list_frame.pack(fill="both", expand=True, pady=(0, 10))

        canvas = tk.Canvas(list_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for i, path in enumerate(self.sample.files):
            ttk.Checkbutton(inner, text=path.name, variable=self.file_vars[path]).grid(row=i, column=0, sticky="w", padx=4, pady=3)

        button_row_bottom = ttk.Frame(outer)
        button_row_bottom.pack(fill="x")
        ttk.Button(button_row_bottom, text="OK", command=self._ok).pack(side="right")
        ttk.Button(button_row_bottom, text="Cancel", command=self._cancel).pack(side="right", padx=(0, 8))

    def _select_all(self):
        for var in self.file_vars.values():
            var.set(True)

    def _select_none(self):
        for var in self.file_vars.values():
            var.set(False)

    def _ok(self):
        selected = [path for path, var in self.file_vars.items() if var.get()]
        if not selected:
            messagebox.showwarning("FRAP Analysis", "Select at least one CSV file for this sample.", parent=self)
            return
        self.result = selected
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


class SampleRow:
    def __init__(self, parent, sample: SampleInput, row_index: int, root_window: tk.Tk):
        self.sample = sample
        self.root_window = root_window
        self.selected_files: List[Path] = list(sample.files)  # Default: analyze all CSV files.
        self.fit_mode_var = tk.StringVar(value="auto")
        self.time_value_var = tk.StringVar(value="90")
        self.time_unit_var = tk.StringVar(value="seconds")

        self.name_label = ttk.Label(parent, text=sample.name)
        self.count_label = ttk.Label(parent, text=self._count_text())
        self.choose_files_button = ttk.Button(parent, text="Choose CSV files", command=self.choose_files)
        self.fit_combo = ttk.Combobox(parent, textvariable=self.fit_mode_var, state="readonly", values=["auto", "fixed"], width=10)
        self.time_entry = ttk.Entry(parent, textvariable=self.time_value_var, width=10)
        self.unit_combo = ttk.Combobox(parent, textvariable=self.time_unit_var, state="readonly", values=["seconds", "minutes"], width=10)
        self.output_label = ttk.Label(parent, text=sample.output_excel.name)

        self.name_label.grid(row=row_index, column=0, sticky="w", padx=4, pady=4)
        self.count_label.grid(row=row_index, column=1, sticky="w", padx=4, pady=4)
        self.choose_files_button.grid(row=row_index, column=2, sticky="w", padx=4, pady=4)
        self.fit_combo.grid(row=row_index, column=3, sticky="w", padx=4, pady=4)
        self.time_entry.grid(row=row_index, column=4, sticky="w", padx=4, pady=4)
        self.unit_combo.grid(row=row_index, column=5, sticky="w", padx=4, pady=4)
        self.output_label.grid(row=row_index, column=6, sticky="w", padx=4, pady=4)

        self.fit_combo.bind("<<ComboboxSelected>>", lambda event: self.update_state())
        self.update_state()

    def _count_text(self) -> str:
        return f"{len(self.selected_files)} / {self.sample.total_csv_count}"

    def choose_files(self):
        dialog = FileSelectionDialog(self.root_window, self.sample, self.selected_files)
        self.root_window.wait_window(dialog)
        if dialog.result is None:
            return
        self.selected_files = list(dialog.result)
        self.count_label.configure(text=self._count_text())

    def update_state(self):
        use_fixed = self.fit_mode_var.get() == "fixed"
        self.time_entry.configure(state="normal" if use_fixed else "disabled")
        self.unit_combo.configure(state="readonly" if use_fixed else "disabled")

    def fixed_time_seconds(self) -> Optional[float]:
        if self.fit_mode_var.get() != "fixed":
            return None

        raw_value = self.time_value_var.get().strip()
        if not raw_value:
            raise ValueError(f"Please enter a timeframe value for sample '{self.sample.name}'.")

        value = float(raw_value)
        if value <= 0:
            raise ValueError(f"Timeframe must be greater than 0 for sample '{self.sample.name}'.")

        if self.time_unit_var.get() == "minutes":
            return value * 60.0
        return value

    def to_config(self) -> SampleRunConfig:
        if not self.selected_files:
            raise ValueError(f"No CSV files selected for sample '{self.sample.name}'.")

        selected_sample = SampleInput(
            name=self.sample.name,
            folder=self.sample.folder,
            files=sorted(self.selected_files),
            output_excel=self.sample.output_excel,
        )
        return SampleRunConfig(
            sample=selected_sample,
            fit_mode=self.fit_mode_var.get(),
            fixed_fit_window_s=self.fixed_time_seconds(),
        )


class FrameFinderTab(ttk.Frame):
    """Find the acquired frame nearest to requested FRAP timepoints."""

    def __init__(self, parent):
        super().__init__(parent, padding=14)
        self.timing: Optional[FRAPFrameTimingData] = None
        self.results: List[FRAPFrameMatch] = []
        self.reference_label = ""

        self.file_label_var = tk.StringVar(value="No raw FRAP CSV selected.")
        self.file_summary_var = tk.StringVar(value="")
        self.reference_mode_var = tk.StringVar(value="bleach")
        self.requested_times_var = tk.StringVar(value="0, 5, 100")
        self.first_frame_number_var = tk.StringVar(value="1")
        self.status_var = tk.StringVar(value="Load one raw FRAP CSV file.")

        self._build_ui()

    def _build_ui(self):
        self.columnconfigure(0, weight=0, minsize=365)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        controls = ttk.Frame(self)
        controls.grid(row=0, column=0, sticky="nsew", padx=(0, 12))

        results_panel = ttk.Frame(self)
        results_panel.grid(row=0, column=1, sticky="nsew")
        results_panel.columnconfigure(0, weight=1)
        results_panel.rowconfigure(2, weight=1)

        ttk.Label(controls, text="Frame Finder", font=("Arial", 14, "bold")).pack(anchor="w", pady=(0, 10))

        file_box = ttk.LabelFrame(controls, text="Raw FRAP file", padding=10)
        file_box.pack(fill="x", pady=(0, 10))
        ttk.Button(file_box, text="Choose CSV file", command=self._choose_file).pack(fill="x")
        ttk.Label(file_box, textvariable=self.file_label_var, wraplength=330, justify="left").pack(fill="x", pady=(7, 0))
        ttk.Label(file_box, textvariable=self.file_summary_var, wraplength=330, justify="left", foreground="#444444").pack(fill="x", pady=(4, 0))

        reference_box = ttk.LabelFrame(controls, text="Time reference", padding=10)
        reference_box.pack(fill="x", pady=(0, 10))
        ttk.Radiobutton(
            reference_box,
            text="From bleaching moment",
            variable=self.reference_mode_var,
            value="bleach",
        ).pack(anchor="w")
        ttk.Radiobutton(
            reference_box,
            text="From beginning of measurement",
            variable=self.reference_mode_var,
            value="beginning",
        ).pack(anchor="w", pady=(4, 0))
        ttk.Label(
            reference_box,
            text="Bleaching is read directly from the Markers column. The bleach-marked frame is 0 s.",
            wraplength=325,
            foreground="#555555",
        ).pack(fill="x", pady=(7, 0))

        request_box = ttk.LabelFrame(controls, text="Requested timepoints", padding=10)
        request_box.pack(fill="x", pady=(0, 10))
        ttk.Label(request_box, text="Times in seconds, separated by commas, spaces, or semicolons:").pack(anchor="w")
        ttk.Entry(request_box, textvariable=self.requested_times_var).pack(fill="x", pady=(5, 8))

        frame_row = ttk.Frame(request_box)
        frame_row.pack(fill="x")
        ttk.Label(frame_row, text="First acquired frame number:").pack(side="left")
        ttk.Spinbox(frame_row, from_=0, to=1000000, textvariable=self.first_frame_number_var, width=9).pack(side="right")

        ttk.Button(controls, text="Find nearest frames", command=self._calculate).pack(fill="x", pady=(0, 6))
        ttk.Button(controls, text="Copy results", command=self._copy_results).pack(fill="x", pady=(0, 6))
        ttk.Button(controls, text="Export result CSV", command=self._export_results).pack(fill="x")

        ttk.Label(
            controls,
            text=(
                "The nearest acquired frame is returned. The exact time of that frame and the signed timing "
                "difference are shown, so a non-exact match is explicit."
            ),
            wraplength=335,
            foreground="#555555",
        ).pack(fill="x", pady=(12, 0))

        ttk.Label(results_panel, text="Matched frames", font=("Arial", 11, "bold")).grid(row=0, column=0, sticky="w")
        self.reference_info_label = ttk.Label(results_panel, text="Reference: not calculated", foreground="#333333")
        self.reference_info_label.grid(row=1, column=0, sticky="w", pady=(4, 8))

        table_frame = ttk.Frame(results_panel)
        table_frame.grid(row=2, column=0, sticky="nsew")
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        columns = (
            "requested",
            "frame",
            "exact_reference",
            "difference",
            "relation",
            "measurement_start",
            "raw_time",
        )
        self.results_tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        headings = {
            "requested": "Requested time (s)",
            "frame": "Frame number",
            "exact_reference": "Exact time from reference (s)",
            "difference": "Difference (s)",
            "relation": "Match position",
            "measurement_start": "Time from measurement start (s)",
            "raw_time": "Raw CSV time (s)",
        }
        widths = {
            "requested": 125,
            "frame": 105,
            "exact_reference": 180,
            "difference": 115,
            "relation": 135,
            "measurement_start": 190,
            "raw_time": 125,
        }
        for column in columns:
            self.results_tree.heading(column, text=headings[column])
            self.results_tree.column(column, width=widths[column], anchor="center", stretch=True)

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.results_tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.results_tree.xview)
        self.results_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.results_tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        status = ttk.LabelFrame(results_panel, text="Status", padding=8)
        status.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(status, textvariable=self.status_var, wraplength=700, justify="left").pack(fill="x")

    def _clear_results(self):
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        self.results = []
        self.reference_label = ""
        self.reference_info_label.configure(text="Reference: not calculated")

    def _choose_file(self):
        filename = filedialog.askopenfilename(
            title="Choose a raw FRAP CSV output",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not filename:
            return

        try:
            timing = parse_frap_frame_timing(Path(filename))
        except Exception as exc:
            messagebox.showerror("Frame Finder", str(exc))
            self.status_var.set(f"Could not load file: {exc}")
            return

        self.timing = timing
        self._clear_results()
        self.file_label_var.set(str(timing.path))

        if timing.bleach_index is None:
            bleach_text = "No bleaching marker found"
        else:
            idx = timing.bleach_index
            bleach_text = (
                f"Bleach frame: {int(timing.frame_numbers[idx])}; "
                f"time from start: {format_time_value(timing.time_from_start_seconds[idx])} s"
            )

        self.file_summary_var.set(
            f"Frames: {len(timing.frame_numbers)}\n"
            f"Source time unit: {timing.source_unit or 'not specified'}\n"
            f"Duration: {format_time_value(timing.time_from_start_seconds[-1])} s\n"
            f"{bleach_text}"
        )
        self.status_var.set("File loaded. Enter requested timepoints and calculate.")

    def _calculate(self):
        if self.timing is None:
            messagebox.showwarning("Frame Finder", "Choose a raw FRAP CSV file first.")
            return

        try:
            requested = parse_requested_timepoints(self.requested_times_var.get())
            first_frame_number = int(self.first_frame_number_var.get().strip())
            results, reference_label = match_frap_timepoints(
                timing=self.timing,
                requested_times_s=requested,
                reference_mode=self.reference_mode_var.get(),
                first_frame_number=first_frame_number,
            )
        except Exception as exc:
            messagebox.showerror("Frame Finder", str(exc))
            self.status_var.set(f"Calculation failed: {exc}")
            return

        self._clear_results()
        self.results = results
        self.reference_label = reference_label
        self.reference_info_label.configure(text=f"Reference: {reference_label}")

        for result in results:
            self.results_tree.insert(
                "",
                "end",
                values=(
                    format_time_value(result.requested_time_s),
                    result.frame_number,
                    format_time_value(result.exact_time_from_reference_s),
                    format_time_value(result.difference_s),
                    result.relation,
                    format_time_value(result.time_from_measurement_start_s),
                    format_time_value(result.raw_csv_time_s),
                ),
            )

        non_exact = sum(result.relation != "Exact" for result in results)
        self.status_var.set(
            f"Matched {len(results)} requested timepoint(s). "
            f"Non-exact or out-of-range matches: {non_exact}."
        )

    def _copy_results(self):
        if not self.results:
            messagebox.showwarning("Frame Finder", "Calculate matching frames first.")
            return

        lines = [
            "requested_time_s\tframe_number\texact_time_from_reference_s\tdifference_s\tmatch_position"
        ]
        for result in self.results:
            lines.append(
                "\t".join(
                    [
                        format_time_value(result.requested_time_s),
                        str(result.frame_number),
                        format_time_value(result.exact_time_from_reference_s),
                        format_time_value(result.difference_s),
                        result.relation,
                    ]
                )
            )
        self.clipboard_clear()
        self.clipboard_append("\n".join(lines))
        self.status_var.set("Results copied to the clipboard.")

    def _export_results(self):
        if not self.results or self.timing is None:
            messagebox.showwarning("Frame Finder", "Calculate matching frames first.")
            return

        output = filedialog.asksaveasfilename(
            title="Save frame matches",
            defaultextension=".csv",
            initialfile=f"{self.timing.path.stem}_frame_matches.csv",
            filetypes=[("CSV files", "*.csv")],
        )
        if not output:
            return

        rows = []
        for result in self.results:
            rows.append(
                {
                    "source_file": self.timing.path.name,
                    "reference": self.reference_label,
                    "requested_time_s": result.requested_time_s,
                    "frame_number": result.frame_number,
                    "exact_time_from_reference_s": result.exact_time_from_reference_s,
                    "time_from_measurement_start_s": result.time_from_measurement_start_s,
                    "raw_csv_time_s": result.raw_csv_time_s,
                    "difference_from_requested_s": result.difference_s,
                    "match_position": result.relation,
                }
            )

        try:
            pd.DataFrame(rows).to_csv(output, index=False, encoding="utf-8-sig")
        except Exception as exc:
            messagebox.showerror("Frame Finder", f"Export failed: {exc}")
            self.status_var.set(f"Export failed: {exc}")
            return

        self.status_var.set(f"Results exported to: {output}")
        messagebox.showinfo("Frame Finder", f"Saved result CSV:\n{output}")


class FRAPGui(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("FRAP Analysis")
        self.geometry("1180x760")
        self.minsize(980, 650)
        self.resizable(True, True)

        self.selected_folder: Optional[Path] = None
        self.sample_rows: List[SampleRow] = []
        self.channel_var = tk.StringVar(value="auto")
        self.open_outputs_var = tk.BooleanVar(value=False)

        self._build_ui()

    def _build_ui(self):
        outer = ttk.Frame(self, padding=10)
        outer.pack(fill="both", expand=True)

        notebook = ttk.Notebook(outer)
        notebook.pack(fill="both", expand=True)

        analysis_tab = ttk.Frame(notebook)
        frame_finder_tab = ttk.Frame(notebook)
        notebook.add(analysis_tab, text="FRAP Analysis")
        notebook.add(frame_finder_tab, text="Frame Finder")

        self._build_analysis_tab(analysis_tab)
        self.frame_finder_tab = FrameFinderTab(frame_finder_tab)
        self.frame_finder_tab.pack(fill="both", expand=True)

    def _build_analysis_tab(self, parent):
        outer = ttk.Frame(parent, padding=16)
        outer.pack(fill="both", expand=True)

        title = ttk.Label(outer, text="FRAP Analysis", font=("Arial", 14, "bold"))
        title.pack(anchor="w", pady=(0, 12))

        folder_frame = ttk.LabelFrame(outer, text="Input folder", padding=12)
        folder_frame.pack(fill="x", pady=(0, 12))

        self.folder_label = ttk.Label(folder_frame, text="No folder selected.", justify="left")
        self.folder_label.pack(side="left", fill="x", expand=True)
        ttk.Button(folder_frame, text="Choose folder", command=self._choose_folder).pack(side="right", padx=(8, 0))

        options = ttk.LabelFrame(outer, text="Global options", padding=12)
        options.pack(fill="x", pady=(0, 12))

        ttk.Label(options, text="Channels to analyze:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        channel_combo = ttk.Combobox(options, textvariable=self.channel_var, state="readonly", values=["auto", "1", "2", "3", "4"], width=12)
        channel_combo.grid(row=0, column=1, sticky="w")

        open_check = ttk.Checkbutton(options, text="Open output file(s) after analysis", variable=self.open_outputs_var)
        open_check.grid(row=0, column=2, sticky="w", padx=(24, 0))

        samples_frame = ttk.LabelFrame(outer, text="Detected samples and fitting timeframe", padding=12)
        samples_frame.pack(fill="both", expand=True, pady=(0, 12))

        self.samples_canvas = tk.Canvas(samples_frame, highlightthickness=0)
        self.samples_scrollbar = ttk.Scrollbar(samples_frame, orient="vertical", command=self.samples_canvas.yview)
        self.samples_inner = ttk.Frame(self.samples_canvas)
        self.samples_inner.bind(
            "<Configure>",
            lambda event: self.samples_canvas.configure(scrollregion=self.samples_canvas.bbox("all")),
        )
        self.samples_canvas.create_window((0, 0), window=self.samples_inner, anchor="nw")
        self.samples_canvas.configure(yscrollcommand=self.samples_scrollbar.set)
        self.samples_canvas.pack(side="left", fill="both", expand=True)
        self.samples_scrollbar.pack(side="right", fill="y")

        self._draw_empty_samples_message()

        button_row = ttk.Frame(outer)
        button_row.pack(fill="x", pady=(0, 12))
        ttk.Button(button_row, text="Run analysis", command=self._run_analysis).pack(side="left")
        ttk.Button(button_row, text="Close", command=self.destroy).pack(side="left", padx=(8, 0))

        status_frame = ttk.LabelFrame(outer, text="Status", padding=12)
        status_frame.pack(fill="both", expand=True)

        self.status_box = tk.Text(status_frame, height=9, wrap="word")
        self.status_box.pack(fill="both", expand=True)
        self.status_box.insert("1.0", "Ready. Choose an input folder.\n")
        self.status_box.configure(state="disabled")

    def _draw_empty_samples_message(self):
        for child in self.samples_inner.winfo_children():
            child.destroy()
        self.sample_rows.clear()
        ttk.Label(self.samples_inner, text="No samples detected yet. Click 'Choose folder'.").grid(row=0, column=0, sticky="w", padx=4, pady=4)

    def _draw_samples_table(self, samples: List[SampleInput]):
        for child in self.samples_inner.winfo_children():
            child.destroy()
        self.sample_rows.clear()

        headers = ["Sample", "CSV selected", "Files", "Fit mode", "Time", "Unit", "Output workbook"]
        for col, header in enumerate(headers):
            ttk.Label(self.samples_inner, text=header, font=("Arial", 9, "bold")).grid(row=0, column=col, sticky="w", padx=4, pady=(0, 6))

        for i, sample in enumerate(samples, start=1):
            self.sample_rows.append(SampleRow(self.samples_inner, sample, i, self))

    def _append_status(self, text: str):
        self.status_box.configure(state="normal")
        self.status_box.insert("end", text + "\n")
        self.status_box.see("end")
        self.status_box.configure(state="disabled")
        self.update_idletasks()

    def _choose_folder(self):
        folder = filedialog.askdirectory(title="Choose folder containing FRAP CSV files or sample subfolders")
        if not folder:
            return

        try:
            chosen_folder = Path(folder).resolve()
            samples = detect_samples(chosen_folder)
            if not samples:
                self.selected_folder = chosen_folder
                self.folder_label.configure(text=f"Selected folder: {chosen_folder}\nNo CSV samples detected.")
                self._draw_empty_samples_message()
                self._append_status(f"No CSV files were found in: {chosen_folder}")
                return

            self.selected_folder = chosen_folder
            self.folder_label.configure(
                text=(
                    f"Selected folder: {chosen_folder}\n"
                    f"Detected samples: {len(samples)}\n"
                    f"Output files will be saved in this selected folder."
                )
            )
            self._draw_samples_table(samples)
            self._append_status(f"Detected {len(samples)} sample(s) in: {chosen_folder}")
            for sample in samples:
                self._append_status(f"  - {sample.name}: {len(sample.files)} CSV file(s)")
        except Exception as exc:
            self._append_status(f"Error while scanning folder: {exc}")
            messagebox.showerror("FRAP Analysis", str(exc))

    def _run_analysis(self):
        if not self.sample_rows:
            messagebox.showwarning("FRAP Analysis", "Choose a folder with CSV files before running analysis.")
            return

        output_paths: List[Path] = []
        try:
            configs = [row.to_config() for row in self.sample_rows]
            self._append_status("Starting analysis...")

            for config in configs:
                self._append_status(f"Analyzing sample: {config.sample.name}")
                result_text = run_analysis_for_sample(
                    sample=config.sample,
                    channel_selection=self.channel_var.get(),
                    fit_mode=config.fit_mode,
                    fixed_fit_window_s=config.fixed_fit_window_s,
                )
                output_paths.append(config.sample.output_excel)
                self._append_status(result_text)

            finished_text = f"Analysis finished successfully. Output workbooks created: {len(output_paths)}"
            self._append_status(finished_text)

            if self.open_outputs_var.get():
                for path in output_paths:
                    open_file_with_default_app(path)

            messagebox.showinfo("FRAP Analysis", finished_text)
        except Exception as exc:
            error_text = f"Error: {exc}\n\n{traceback.format_exc()}"
            self._append_status(error_text)
            messagebox.showerror("FRAP Analysis", str(exc))


if __name__ == "__main__":
    app = FRAPGui()
    try:
        style = ttk.Style(app)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass
    app.mainloop()
