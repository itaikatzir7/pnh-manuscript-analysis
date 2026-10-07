"""Trypsin Assay Analyzer 3.2

Source filename: trypsin_assay_gui_v4.py; original internal version 3.2 retained.
Run: python trypsin_assay_gui_v4.py
Dependencies: Python 3.10+, numpy, matplotlib, xlsxwriter; tkinter for the GUI.
Install dependencies: python -m pip install numpy matplotlib xlsxwriter

CLI example:
  python trypsin_assay_gui_v4.py --input source.xlsx --output analysis.xlsx

Input: the same 18-well worksheet layout as v3.1 (three technical wells each
for oligoT control/trypsin, oligoC control/trypsin, oligoA control/trypsin).
Each worksheet is one system and independent experimental iteration.
The time origin is the first recorded measurement, as in v3.1.

No normalization is performed. Plateau thresholds are operational settings,
not a universal biochemical definition. Inspect the shaded fit intervals in
Review curves and windows. Manual endpoints are recorded in the export.
Dissolution rate = signed control slope minus signed trypsin slope.
Exported means and sample SDs use independent iterations. Rates are always
in AU/min. All recorded data are displayed; no plotting limit truncates fits.
"""
from __future__ import annotations

import argparse
import math
import os
import re
import tempfile
import threading
import traceback
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Optional
from xml.etree import ElementTree as ET

import numpy as np
import xlsxwriter
from matplotlib.figure import Figure

SOFTWARE_VERSION = "Trypsin Assay Analyzer v3.2 - raw dissolution rates"


GROUP_SPECS = [
    ("oligoT", "Control"),
    ("oligoT", "Trypsin"),
    ("oligoC", "Control"),
    ("oligoC", "Trypsin"),
    ("oligoA", "Control"),
    ("oligoA", "Trypsin"),
]

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _excel_col_to_index(cell_ref: str) -> int:
    match = re.match(r"([A-Za-z]+)", cell_ref)
    if not match:
        raise ValueError(f"Invalid Excel cell reference: {cell_ref!r}")
    value = 0
    for char in match.group(1).upper():
        value = value * 26 + (ord(char) - 64)
    return value - 1


def _excel_index_to_col(index: int) -> str:
    index += 1
    letters: list[str] = []
    while index:
        index, remainder = divmod(index - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def _safe_float(value: Any) -> float:
    if value is None or value == "":
        return math.nan
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float, np.number)):
        return float(value)
    text = str(value).strip().replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return math.nan


def _parse_time_string_to_minutes(text: str) -> Optional[float]:
    text = text.strip()
    if not text:
        return None
    parts = text.split(":")
    try:
        if len(parts) == 3:
            hours, minutes, seconds = map(float, parts)
            return hours * 60.0 + minutes + seconds / 60.0
        if len(parts) == 2:
            minutes, seconds = map(float, parts)
            return minutes + seconds / 60.0
    except ValueError:
        return None
    return None


def _time_raw_scalar(value: Any) -> tuple[Optional[float], bool]:
    """Return scalar and whether it was parsed from an hh:mm:ss-like string."""
    if value is None or value == "":
        return None, False
    if isinstance(value, (int, float, np.number)) and not isinstance(value, bool):
        scalar = float(value)
        if math.isfinite(scalar):
            return scalar, False
        return None, False
    parsed = _parse_time_string_to_minutes(str(value))
    if parsed is not None:
        return parsed, True
    try:
        scalar = float(str(value).strip())
        if math.isfinite(scalar):
            return scalar, False
    except ValueError:
        pass
    return None, False


def _is_all_zero_or_blank(values: Iterable[Any], tolerance: float = 1e-12) -> bool:
    saw_numeric = False
    for value in values:
        if value is None or value == "":
            continue
        number = _safe_float(value)
        if math.isnan(number):
            return False
        saw_numeric = True
        if abs(number) > tolerance:
            return False
    return saw_numeric


def _nanstd_triplicates(values: np.ndarray) -> np.ndarray:
    result = np.full(values.shape[0], np.nan, dtype=float)
    for row_idx in range(values.shape[0]):
        row = values[row_idx, :]
        finite = row[np.isfinite(row)]
        if finite.size >= 2:
            result[row_idx] = float(np.std(finite, ddof=1))
        elif finite.size == 1:
            result[row_idx] = 0.0
    return result


def _baseline(values: np.ndarray, points: int) -> float:
    points = max(1, int(points))
    finite = values[:points][np.isfinite(values[:points])]
    if finite.size == 0:
        return math.nan
    return float(np.mean(finite))


def _parse_system_iteration(sheet_name: str) -> tuple[str, Optional[int]]:
    match = re.match(r"^(.*?)[_\-\s]+(\d+)\s*$", sheet_name.strip())
    if not match:
        return sheet_name.strip(), None
    return match.group(1).strip(), int(match.group(2))


def _extract_wavelength(header: Any) -> Optional[int]:
    if header is None:
        return None
    matches = re.findall(r"\d{3,4}", str(header))
    if not matches:
        return None
    return int(matches[-1])


class XlsxRawReader:
    """Minimal XLSX reader for values, avoiding a dependency on Excel automation."""

    def __init__(self, path: str | Path):
        self.path = Path(path)
        if not self.path.exists():
            raise FileNotFoundError(self.path)
        self._shared_strings: list[str] = []
        self._sheet_paths: dict[str, str] = {}
        self._load_metadata()

    @property
    def sheet_names(self) -> list[str]:
        return list(self._sheet_paths)

    def _load_metadata(self) -> None:
        with zipfile.ZipFile(self.path) as archive:
            if "xl/sharedStrings.xml" in archive.namelist():
                root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
                for item in root:
                    if _local_name(item.tag) != "si":
                        continue
                    text = "".join(node.text or "" for node in item.iter() if _local_name(node.tag) == "t")
                    self._shared_strings.append(text)

            workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
            relationships_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            relationship_targets: dict[str, str] = {}
            for relation in relationships_root:
                if _local_name(relation.tag) == "Relationship":
                    relationship_targets[relation.attrib["Id"]] = relation.attrib["Target"]

            for sheet in workbook_root.iter():
                if _local_name(sheet.tag) != "sheet":
                    continue
                name = sheet.attrib["name"]
                rel_id = sheet.attrib.get(f"{{{REL_NS}}}id")
                if not rel_id or rel_id not in relationship_targets:
                    continue
                target = relationship_targets[rel_id].replace("\\", "/")
                if target.startswith("/"):
                    target = target.lstrip("/")
                elif not target.startswith("xl/"):
                    target = f"xl/{target}"
                target = os.path.normpath(target).replace("\\", "/")
                self._sheet_paths[name] = target

    def read_sheet(self, sheet_name: str) -> list[list[Any]]:
        if sheet_name not in self._sheet_paths:
            raise KeyError(sheet_name)
        with zipfile.ZipFile(self.path) as archive:
            xml_bytes = archive.read(self._sheet_paths[sheet_name])
        root = ET.fromstring(xml_bytes)
        row_maps: list[dict[int, Any]] = []
        max_col = -1

        for row in root.iter():
            if _local_name(row.tag) != "row":
                continue
            row_values: dict[int, Any] = {}
            for cell in row:
                if _local_name(cell.tag) != "c":
                    continue
                reference = cell.attrib.get("r")
                if not reference:
                    continue
                col_idx = _excel_col_to_index(reference)
                cell_type = cell.attrib.get("t")
                value_node = next((node for node in cell if _local_name(node.tag) == "v"), None)
                inline_node = next((node for node in cell if _local_name(node.tag) == "is"), None)
                value: Any = None

                if cell_type == "inlineStr" and inline_node is not None:
                    value = "".join(node.text or "" for node in inline_node.iter() if _local_name(node.tag) == "t")
                elif value_node is not None and value_node.text is not None:
                    raw = value_node.text
                    if cell_type == "s":
                        try:
                            value = self._shared_strings[int(raw)]
                        except (ValueError, IndexError):
                            value = raw
                    elif cell_type == "b":
                        value = raw == "1"
                    elif cell_type in {"str", "e"}:
                        value = raw
                    else:
                        try:
                            numeric = float(raw)
                            value = int(numeric) if numeric.is_integer() else numeric
                        except ValueError:
                            value = raw
                row_values[col_idx] = value
                max_col = max(max_col, col_idx)
            row_maps.append(row_values)

        width = max_col + 1 if max_col >= 0 else 0
        return [[row_map.get(col_idx) for col_idx in range(width)] for row_map in row_maps]


@dataclass
class GroupData:
    oligo: str
    condition: str
    well_labels: list[str]
    source_columns: list[str]
    values: np.ndarray
    mean: np.ndarray
    std: np.ndarray


@dataclass
class SheetData:
    sheet_name: str
    system: str
    iteration: Optional[int]
    header_row_number: int
    time_minutes: np.ndarray
    temperature_header: str
    wavelength_nm: Optional[int]
    data_headers: list[str]
    source_columns: list[str]
    raw_values: np.ndarray
    groups: list[GroupData]
    warnings: list[str] = field(default_factory=list)

    @property
    def duration_minutes(self) -> float:
        if self.time_minutes.size == 0:
            return 0.0
        return float(self.time_minutes[-1] - self.time_minutes[0])

    def group(self, oligo: str, condition: str) -> GroupData:
        for group in self.groups:
            if group.oligo == oligo and group.condition == condition:
                return group
        raise KeyError((oligo, condition))

class TrypsinWorkbookAnalyzer:
    def __init__(self, path: str | Path):
        self.path = Path(path)

    def analyze(self) -> list[SheetData]:
        reader = XlsxRawReader(self.path)
        analyses: list[SheetData] = []
        errors: list[str] = []
        for sheet_name in reader.sheet_names:
            try:
                analyses.append(self._analyze_sheet(sheet_name, reader.read_sheet(sheet_name)))
            except Exception as exc:
                errors.append(f"{sheet_name}: {exc}")
        if errors:
            raise ValueError("Some worksheets could not be analyzed:\n" + "\n".join(errors))
        if not analyses:
            raise ValueError("No analyzable worksheets were found.")
        return analyses

    def _analyze_sheet(self, sheet_name: str, rows: list[list[Any]]) -> SheetData:
        header_row_idx, time_col_idx, temp_col_idx, data_col_indices = self._find_header(rows)
        header = rows[header_row_idx]
        data_headers = [str(header[idx]).strip() for idx in data_col_indices]
        source_columns = [_excel_index_to_col(idx) for idx in data_col_indices]
        if len(data_headers) != 18:
            raise ValueError(
                f"Expected 18 well columns after Time and temperature, but found {len(data_headers)}."
            )

        raw_times: list[float] = []
        time_from_string: list[bool] = []
        matrix_rows: list[list[float]] = []
        warnings: list[str] = []
        previous_time: Optional[float] = None

        for row_idx in range(header_row_idx + 1, len(rows)):
            row = rows[row_idx]
            time_value = row[time_col_idx] if time_col_idx < len(row) else None
            scalar_time, parsed_from_string = _time_raw_scalar(time_value)
            if scalar_time is None:
                if raw_times:
                    break
                continue

            data_values_raw = [row[idx] if idx < len(row) else None for idx in data_col_indices]
            if raw_times and previous_time is not None and scalar_time < previous_time - 1e-10:
                warnings.append(f"Ignored content after a time reset at source row {row_idx + 1}.")
                break
            if raw_times and scalar_time == 0 and previous_time and _is_all_zero_or_blank(data_values_raw):
                warnings.append(f"Ignored trailing zero rows starting at source row {row_idx + 1}.")
                break
            if all(value is None or value == "" for value in data_values_raw):
                if raw_times:
                    break
                continue

            numeric_values = [_safe_float(value) for value in data_values_raw]
            if all(math.isnan(value) for value in numeric_values):
                if raw_times:
                    break
                continue

            raw_times.append(float(scalar_time))
            time_from_string.append(parsed_from_string)
            matrix_rows.append(numeric_values)
            previous_time = float(scalar_time)

        if len(raw_times) < 4:
            raise ValueError("Fewer than four valid measurement rows were found.")

        matrix = np.asarray(matrix_rows, dtype=float)
        zero_rows = np.all(np.isfinite(matrix) & (np.abs(matrix) <= 1e-12), axis=1)
        trim_at: Optional[int] = None
        consecutive = 0
        for idx, is_zero in enumerate(zero_rows):
            consecutive = consecutive + 1 if is_zero else 0
            if consecutive >= 3:
                trim_at = idx - consecutive + 1
                break
        if trim_at is not None and trim_at >= 4:
            matrix = matrix[:trim_at, :]
            raw_times = raw_times[:trim_at]
            time_from_string = time_from_string[:trim_at]
            warnings.append("Ignored three or more consecutive all-zero measurement rows at the end of the data block.")

        time_array = np.asarray(raw_times, dtype=float)
        if any(time_from_string):
            time_minutes = time_array
        else:
            differences = np.diff(time_array)
            positive_differences = differences[differences > 0]
            median_difference = float(np.median(positive_differences)) if positive_differences.size else 0.0
            if float(np.nanmax(time_array)) <= 2.5 and median_difference <= 0.02:
                time_minutes = time_array * 1440.0
            else:
                time_minutes = time_array
        time_minutes = time_minutes - time_minutes[0]

        groups: list[GroupData] = []
        for group_idx, (oligo, condition) in enumerate(GROUP_SPECS):
            start = group_idx * 3
            end = start + 3
            values = matrix[:, start:end]
            groups.append(
                GroupData(
                    oligo=oligo,
                    condition=condition,
                    well_labels=data_headers[start:end],
                    source_columns=source_columns[start:end],
                    values=values,
                    mean=np.nanmean(values, axis=1),
                    std=_nanstd_triplicates(values),
                )
            )

        system, iteration = _parse_system_iteration(sheet_name)
        temperature_header = str(header[temp_col_idx]).strip() if temp_col_idx is not None else ""
        return SheetData(
            sheet_name=sheet_name,
            system=system,
            iteration=iteration,
            header_row_number=header_row_idx + 1,
            time_minutes=time_minutes,
            temperature_header=temperature_header,
            wavelength_nm=_extract_wavelength(temperature_header),
            data_headers=data_headers,
            source_columns=source_columns,
            raw_values=matrix,
            groups=groups,
            warnings=warnings,
        )

    @staticmethod
    def _find_header(rows: list[list[Any]]) -> tuple[int, int, int, list[int]]:
        for row_idx, row in enumerate(rows):
            normalized = [str(value).strip().lower() if value is not None else "" for value in row]
            for time_idx, value in enumerate(normalized):
                if value != "time":
                    continue
                nonempty_after = [idx for idx in range(time_idx + 1, len(row)) if row[idx] not in (None, "")]
                if len(nonempty_after) < 7:
                    continue
                temp_idx = next(
                    (
                        idx
                        for idx in nonempty_after
                        if "temp" in normalized[idx]
                        or "°" in str(row[idx])
                        or normalized[idx].startswith("t°")
                    ),
                    nonempty_after[0],
                )
                data_indices = [idx for idx in nonempty_after if idx != temp_idx]
                return row_idx, time_idx, temp_idx, data_indices
        raise ValueError("Could not find a header row containing 'Time' and well labels.")



@dataclass
class Settings:
    start: float = 5.0
    plateau_duration: float = 10.0
    plateau_absolute: float = 0.005
    plateau_fraction: float = 0.02
    review_r2: float = 0.95
    common_y: bool = True

    def validate(self):
        vals = [self.start, self.plateau_duration, self.plateau_absolute,
                self.plateau_fraction, self.review_r2]
        if not all(math.isfinite(v) for v in vals):
            raise ValueError('Settings must be finite numbers.')
        if self.start < 0 or self.plateau_duration <= 0 or self.plateau_absolute <= 0:
            raise ValueError('Start must be nonnegative; duration and absolute tolerance must be positive.')
        if not 0 <= self.plateau_fraction <= 0.2 or not 0 <= self.review_r2 <= 1:
            raise ValueError('Plateau fraction must be 0 to 0.2; review R² must be 0 to 1.')


def linear_fit(t, y):
    t, y = np.asarray(t), np.asarray(y)
    valid = np.isfinite(t) & np.isfinite(y)
    x, z = t[valid], y[valid]
    if len(x) < 4 or np.ptp(x) <= 0:
        return (math.nan, math.nan, math.nan, len(x))
    m, b = np.polyfit(x, z, 1)
    residual = np.sum((z - (m*x+b))**2)
    total = np.sum((z-z.mean())**2)
    r2 = 1-residual/total if total > 1e-20 else (1.0 if residual < 1e-20 else 0.0)
    return float(m), float(b), float(r2), len(x)


def detect_plateau(t, y, settings):
    """Conservative terminal stability rule, not an enzymatic kinetic model.

    A three-point median suppresses isolated spikes only for detection. A
    plateau is the earliest sustained terminal suffix within tolerance of the
    final-window median, with negligible total fitted drift. Original values
    are always used for rate fits. A slow drift below tolerance is unresolved.
    """
    valid = np.isfinite(t) & np.isfinite(y) & (t >= settings.start-1e-9)
    x, z = np.asarray(t)[valid], np.asarray(y)[valid]
    if len(x) < 4:
        return math.nan, math.nan, 'Insufficient data after start'
    smooth = np.array([np.median(z[max(0,i-1):min(len(z),i+2)]) for i in range(len(z))])
    amplitude = float(np.ptp(smooth))
    tolerance = max(settings.plateau_absolute, settings.plateau_fraction*amplitude)
    tail = x >= x[-1]-settings.plateau_duration-1e-9
    tail_x, tail_z = x[tail], smooth[tail]
    if len(tail_x) < 4 or tail_x[-1]-tail_x[0] < settings.plateau_duration-1e-6:
        return float(x[-1]), tolerance, 'No confirmed plateau: insufficient terminal duration'
    terminal = float(np.median(tail_z))
    if np.max(smooth)-terminal <= 3*tolerance:
        return float(x[-1]), tolerance, 'No resolved decay-to-plateau transition; fitted to end'
    tail_slope = linear_fit(tail_x, tail_z)[0]
    if abs(tail_slope)*(tail_x[-1]-tail_x[0]) > tolerance:
        return float(x[-1]), tolerance, 'No confirmed plateau: terminal drift; fitted to end'
    for i in range(len(x)-3):
        if x[-1]-x[i] < settings.plateau_duration-1e-6:
            break
        suffix = smooth[i:]
        if np.max(np.abs(suffix-terminal)) > tolerance:
            continue
        slope = linear_fit(x[i:], suffix)[0]
        if abs(slope)*(x[-1]-x[i]) > tolerance:
            continue
        if i < 3:
            return float(x[-1]), tolerance, 'Decay completed before usable fit interval; review'
        return float(x[i]), tolerance, 'Terminal plateau detected'
    return float(x[-1]), tolerance, 'No sustained terminal plateau; fitted to end'


@dataclass
class PairResult:
    analysis: SheetData
    oligo: str
    start: float
    end: float
    tolerance: float
    detection: str
    control: tuple
    trypsin: tuple
    control_wells: list
    trypsin_wells: list
    mask: np.ndarray
    review: str

    @property
    def rate(self):
        return self.control[0]-self.trypsin[0]

    @property
    def key(self):
        return self.analysis.sheet_name, self.oligo


def analyze_pair(analysis, oligo, settings, override=None):
    control = analysis.group(oligo, 'Control')
    treated = analysis.group(oligo, 'Trypsin')
    t = analysis.time_minutes
    end, tolerance, status = detect_plateau(t, treated.mean, settings)
    if override is not None:
        if not math.isfinite(override) or not settings.start < override <= t[-1]+1e-9:
            raise ValueError(f'{analysis.sheet_name}/{oligo}: manual end must be after start and within recorded data.')
        end, status = override, 'Manual fit end override'
    mask = ((t >= settings.start-1e-9) & (t <= end+1e-9)
            & np.isfinite(control.mean) & np.isfinite(treated.mean))
    cf, tf = linear_fit(t[mask], control.mean[mask]), linear_fit(t[mask], treated.mean[mask])
    cw = [linear_fit(t[mask], y[mask])[0] for y in control.values.T]
    tw = [linear_fit(t[mask], y[mask])[0] for y in treated.values.T]
    flags = []
    if cf[3] < 4: flags.append('Insufficient matched points')
    if math.isfinite(tf[2]) and tf[2] < settings.review_r2:
        flags.append('Trypsin curve is nonlinear/noisy: slope is an interval average')
    if tf[0] >= 0: flags.append('Trypsin turbidity does not decrease over fit interval')
    if math.isfinite(cf[0]-tf[0]) and cf[0]-tf[0] < 0:
        flags.append('Negative dissolution rate retained')
    return PairResult(analysis, oligo, float(t[mask][0]) if mask.any() else settings.start,
                      float(t[mask][-1]) if mask.any() else end, tolerance, status,
                      cf, tf, cw, tw, mask, '; '.join(flags) or 'OK')


def method_text(s):
    return (
        f'Raw turbidity traces were analyzed without baseline or plateau normalization. '
        f'For each peptide–DNA system and independent experimental iteration, the trypsin-treated '
        f'triplicate-mean trace determined a common fitting interval beginning at {s.start:g} min '
        f'and ending at the onset of a sustained terminal plateau, or at the last valid measurement '
        f'when no plateau was confirmed. Plateau detection used a three-point median-smoothed '
        f'copy only: the terminal suffix had to persist for at least {s.plateau_duration:g} min, '
        f'remain within max({s.plateau_absolute:g} AU, {100*s.plateau_fraction:g}% of the post-start '
        f'trypsin signal range) of the final-window median, and have total fitted drift no greater '
        f'than that tolerance. A preceding decrease exceeding three times the tolerance was required. '
        f'Manual endpoint overrides, if used, are identified in the rate table. Ordinary least-squares '
        f'lines with free intercepts were fitted to the original, unsmoothed trypsin and untreated-control '
        f'mean traces at the same time points. Dissolution rate was calculated as the signed control '
        f'slope minus the signed trypsin slope (AU/min); zero and negative results were retained. '
        f'Technical-well slopes were also fitted over this interval. Each independent iteration '
        f'contributed one corrected rate; reported summary means and sample SDs describe those '
        f'independent iteration-level values, not pooled technical wells. This is an apparent, '
        f'control-corrected turbidity dissolution rate averaged over the selected interval, not a '
        f'direct peptide-cleavage rate or enzyme activity constant. Increasing controls can yield a '
        f'positive correction even when treated turbidity is stable. Slow terminal drift smaller than '
        f'the stated detection tolerance cannot be distinguished from a plateau. Fit quality and '
        f'intervals are reported for review; low R² does not automatically exclude a result.'
    )


def make_figure(result, ylim=None):
    fig = Figure(figsize=(8, 4.6), dpi=110)
    ax = fig.add_subplot(111)
    t = result.analysis.time_minutes
    for condition, fit, color in [('Control', result.control, '#557A8C'), ('Trypsin', result.trypsin, '#AE648A')]:
        g = result.analysis.group(result.oligo, condition)
        ax.plot(t, g.mean, color=color, label=condition, lw=1.8)
        ax.fill_between(t, g.mean-g.std, g.mean+g.std, color=color, alpha=.15)
        x = t[result.mask]
        if len(x) and math.isfinite(fit[0]):
            ax.plot(x, fit[0]*x+fit[1], '--', color=color, lw=1.3)
    ax.axvspan(result.start, result.end, color='grey', alpha=.10)
    ax.axvline(result.end, color='grey', lw=.8, ls=':')
    ax.set(xlabel='Time (min)', ylabel='Turbidity (AU)',
           title=f'{result.analysis.sheet_name} | {result.oligo} | dissolution rate {result.rate:.4g} AU/min')
    if ylim: ax.set_ylim(*ylim)
    ax.set_xlim(left=0, right=max(t[-1], 1))
    ax.legend(frameon=False)
    ax.spines[['top','right']].set_visible(False)
    fig.text(.11, .015, f'Fit: {result.start:g}–{result.end:g} min. {result.detection}', fontsize=8)
    fig.tight_layout(rect=(0,.04,1,1))
    return fig


class Exporter:
    def __init__(self, analyses, settings, source, overrides=None):
        settings.validate()
        self.analyses, self.settings, self.source = analyses, settings, str(source)
        self.overrides = overrides or {}
        self.results = [analyze_pair(a,o,settings,self.overrides.get((a.sheet_name,o)))
                        for a in analyses for o in ('oligoT','oligoC','oligoA')]

    def ylim(self):
        if not self.settings.common_y: return None
        vals = np.concatenate([a.raw_values.ravel() for a in self.analyses])
        vals = vals[np.isfinite(vals)]
        lo, hi = float(vals.min()), float(vals.max())
        pad = max((hi-lo)*.05, .005)
        return lo-pad, hi+pad

    def export(self, output):
        output = Path(output)
        if output.resolve() == Path(self.source).resolve():
            raise ValueError('Choose a different filename to preserve the source workbook.')
        output.parent.mkdir(parents=True, exist_ok=True)
        # Build beside the destination, then atomically replace only on success.
        import io
        fd, temporary = tempfile.mkstemp(suffix='.xlsx', dir=output.parent)
        os.close(fd)
        try:
            with xlsxwriter.Workbook(temporary) as wb:
                head = wb.add_format({'bold':True,'bg_color':'#DCE8ED','text_wrap':True,'border':1})
                wrap = wb.add_format({'text_wrap':True,'valign':'top'})
                number = wb.add_format({'num_format':'0.000000'})
                title = wb.add_format({'bold':True,'font_size':16,'font_color':'#294B60'})
                overview = wb.add_worksheet('Overview')
                overview.set_column('A:H', 16)
                overview.merge_range('A1:H1', SOFTWARE_VERSION, title)
                overview.merge_range('A3:H13', method_text(self.settings), wrap)
                for line in range(2,13): overview.set_row(line,32)
                overview.write_row(14,0,['Setting','Value'],head)
                for row,(key,value) in enumerate(vars(self.settings).items(),15):
                    overview.write(row,0,key); overview.write(row,1,value)
                overview.write(23,0,'Source workbook'); overview.merge_range(23,1,23,7,self.source,wrap)
                overview.write_row(25,0,['Worksheet','System','Iteration','Points','Duration (min)','Import notes'],head)
                for row,a in enumerate(self.analyses,26):
                    overview.write_row(row,0,[a.sheet_name,a.system,a.iteration,len(a.time_minutes),a.duration_minutes,'; '.join(a.warnings)])
                overview.freeze_panes(2,0)
                rates = wb.add_worksheet('Dissolution_Rates')
                headers = ['System','DNA','Iteration','Worksheet','Fit start (min)','Fit end (min)',
                    'Matched points','Control slope (AU/min)','Trypsin slope (AU/min)',
                    'Dissolution rate (AU/min)','Control R²','Trypsin R²','Plateau tolerance (AU)',
                    'Endpoint decision','Review','Control technical SD','Trypsin technical SD']
                rates.write_row(0,0,headers,head); rates.set_row(0,45)
                rates.set_column(0,12,18); rates.set_column(13,14,55); rates.set_column(15,16,22)
                rates.freeze_panes(1,4); rates.autofilter(0,0,len(self.results),len(headers)-1)
                technical = wb.add_worksheet('Technical_Slopes')
                technical.write_row(0,0,['Worksheet','System','DNA','Iteration','Condition','Well',
                    'Signed slope (AU/min)','Fit start (min)','Fit end (min)'],head)
                technical.set_column(0,8,20); technical.freeze_panes(1,0)
                techrow=1
                def safe(ws,row,col,value):
                    if isinstance(value,(float,np.floating)) and not math.isfinite(value): return
                    ws.write(row,col,value,number if isinstance(value,(float,np.floating)) else None)
                grouped={}
                for row,r in enumerate(self.results,1):
                    sd=lambda v: float(np.std(np.array(v)[np.isfinite(v)],ddof=1)) if np.isfinite(v).sum()>1 else math.nan
                    vals=[r.analysis.system,r.oligo,r.analysis.iteration,r.analysis.sheet_name,r.start,r.end,r.control[3],
                          r.control[0],r.trypsin[0],r.rate,r.control[2],r.trypsin[2],r.tolerance,r.detection,r.review,
                          sd(r.control_wells),sd(r.trypsin_wells)]
                    for col,v in enumerate(vals): safe(rates,row,col,v)
                    if math.isfinite(r.rate):
                        rates.write_formula(row,9,f'=H{row+1}-I{row+1}',number,r.rate)
                    grouped.setdefault((r.analysis.system,r.oligo),[]).append((row,r))
                    for condition,slopes in [('Control',r.control_wells),('Trypsin',r.trypsin_wells)]:
                        for well,slope in zip(r.analysis.group(r.oligo,condition).well_labels,slopes):
                            for col,v in enumerate([r.analysis.sheet_name,r.analysis.system,r.oligo,r.analysis.iteration,
                                                   condition,well,slope,r.start,r.end]):safe(technical,techrow,col,v)
                            techrow+=1
                summary=wb.add_worksheet('Summary'); summary.set_column(0,5,23)
                summary.write_row(0,0,['System','DNA','Independent iterations','Mean dissolution rate (AU/min)',
                                       'SD across iterations','Individual day values follow'],head)
                summary.set_row(0,45); summary.freeze_panes(1,2)
                max_days=max(len(v) for v in grouped.values())
                for col in range(max_days):summary.write(0,5+col,f'Iteration value {col+1} (AU/min)',head)
                summary.set_column(5,5+max_days,23)
                labels=[]; plotvalues=[]
                for row,((system,oligo),entries) in enumerate(grouped.items(),1):
                    valid=[(i,r) for i,r in entries if math.isfinite(r.rate)]
                    values=[r.rate for _,r in valid]
                    refs=','.join(f"Dissolution_Rates!J{i+1}" for i,_ in valid)
                    summary.write_row(row,0,[system,oligo,len(valid)])
                    if valid:
                        summary.write_formula(row,3,f'=AVERAGE({refs})',number,float(np.mean(values)))
                    if len(valid)>1:
                        summary.write_formula(row,4,f'=STDEV({refs})',number,float(np.std(values,ddof=1)))
                    for col,(i,r) in enumerate(entries,5):
                        if math.isfinite(r.rate):summary.write_formula(row,col,f'=Dissolution_Rates!J{i+1}',number,r.rate)
                    labels.append(system+'-'+oligo.replace('oligo','o')); plotvalues.append(values)
                fig=Figure(figsize=(12,5),dpi=130); ax=fig.add_subplot(111)
                for i,v in enumerate(plotvalues):
                    if not v: continue
                    ax.scatter(i+np.linspace(-.10,.10,len(v)),v,color='#557A8C',s=22,zorder=3)
                    ax.errorbar(i,np.mean(v),yerr=np.std(v,ddof=1) if len(v)>1 else 0,
                                fmt='_',color='black',capsize=4,markersize=15)
                ax.axhline(0,color='grey',lw=.8); ax.set_xticks(range(len(labels)),labels,rotation=60,ha='right')
                ax.set_ylabel('Dissolution rate (AU/min)'); ax.set_title('Independent iterations; mean ± SD')
                ax.spines[['top','right']].set_visible(False); fig.tight_layout()
                buf=io.BytesIO(); fig.savefig(buf,format='png'); buf.seek(0)
                summary.insert_image(len(grouped)+3,0,'summary.png',{'image_data':buf})
                raw=wb.add_worksheet('Raw_Data'); raw.freeze_panes(1,0)
                rawheaders=['Worksheet','System','Iteration','Time (min)']+[f'{o}_{c}_{i}' for o,c in GROUP_SPECS for i in (1,2,3)]
                raw.write_row(0,0,rawheaders,head); raw.set_column(0,21,18)
                rr=1
                for a in self.analyses:
                    for t,values in zip(a.time_minutes,a.raw_values):
                        for col,v in enumerate([a.sheet_name,a.system,a.iteration,float(t),*values]):safe(raw,rr,col,v)
                        rr+=1
                raw.autofilter(0,0,rr-1,21)
                plots=wb.add_worksheet('Plots'); plots.set_column(0,24,12)
                for i,r in enumerate(self.results):
                    fig=make_figure(r,self.ylim()); buf=io.BytesIO(); fig.savefig(buf,format='png',dpi=115); buf.seek(0)
                    row=(i//2)*29; col=(i%2)*10
                    plots.insert_image(row,col,'curve.png',{'image_data':buf,'x_scale':.8,'y_scale':.8})
                # Put final outputs directly after the explanation.
                wb.worksheets_objs.sort(key=lambda ws:['Overview','Summary','Dissolution_Rates','Technical_Slopes','Raw_Data','Plots'].index(ws.name))
            os.replace(temporary,output)
        finally:
            if os.path.exists(temporary):os.unlink(temporary)


class TrypsinAssayApp:
    def __init__(self):
        import tkinter as tk
        from tkinter import ttk
        self.tk,self.ttk=tk,ttk
        self.root=tk.Tk(); self.root.title(SOFTWARE_VERSION); self.root.geometry('1180x850')
        self.analyses=[]; self.overrides={}; self.canvas=None
        self.input=tk.StringVar(); self.output=tk.StringVar(); self.status=tk.StringVar(value='Select a source workbook.')
        top=ttk.Frame(self.root,padding=10); top.pack(fill='x')
        for row,(label,var,command) in enumerate([('Input workbook',self.input,self.browse_input),('Output workbook',self.output,self.browse_output)]):
            ttk.Label(top,text=label).grid(row=row,column=0,sticky='w')
            ttk.Entry(top,textvariable=var,width=105).grid(row=row,column=1,padx=6,pady=4)
            ttk.Button(top,text='Browse',command=command).grid(row=row,column=2)
        ttk.Button(top,text='Load data',command=self.load).grid(row=2,column=1,sticky='w')
        notebook=ttk.Notebook(self.root); notebook.pack(fill='both',expand=True,padx=10)
        first=ttk.Frame(notebook,padding=12); preview=ttk.Frame(notebook,padding=10)
        notebook.add(first,text='Analysis and export'); notebook.add(preview,text='Review curves and windows')
        settings=Settings(); self.variables={}
        labels={'start':'Fit start (min)','plateau_duration':'Minimum terminal plateau duration (min)',
                'plateau_absolute':'Absolute plateau tolerance (AU)','plateau_fraction':'Plateau tolerance fraction of signal range (0.02 = 2%)',
                'review_r2':'Trypsin R² review threshold (does not exclude results)'}
        for row,(key,label) in enumerate(labels.items()):
            self.variables[key]=tk.StringVar(value=str(getattr(settings,key)))
            ttk.Label(first,text=label).grid(row=row,column=0,sticky='w',pady=5)
            ttk.Entry(first,textvariable=self.variables[key],width=14).grid(row=row,column=1,sticky='w')
        self.common=tk.BooleanVar(value=True)
        ttk.Checkbutton(first,text='Common Y-axis scale',variable=self.common).grid(row=5,column=0,sticky='w')
        ttk.Label(first,text='Dissolution rate = signed control slope − signed trypsin slope (AU/min).\n'
                  'Raw data only. Each matched pair uses the trypsin-selected interval.\n'
                  'Curved traces yield an interval-average rate. Review detected endpoints before export.',
                  wraplength=1000,justify='left').grid(row=6,column=0,columnspan=3,sticky='w',pady=20)
        self.export_button=ttk.Button(first,text='Analyze and export',command=self.export)
        self.export_button.grid(row=7,column=0,sticky='w')
        controls=ttk.Frame(preview); controls.pack(fill='x')
        self.selection=tk.StringVar(); self.selector=ttk.Combobox(controls,textvariable=self.selection,state='readonly',width=38)
        self.selector.pack(side='left'); self.selector.bind('<<ComboboxSelected>>',self.refresh)
        ttk.Button(controls,text='Refresh',command=self.refresh).pack(side='left',padx=5)
        self.manual=tk.StringVar()
        ttk.Label(controls,text='Manual end (min):').pack(side='left',padx=6)
        ttk.Entry(controls,textvariable=self.manual,width=8).pack(side='left')
        ttk.Button(controls,text='Apply',command=self.apply_override).pack(side='left')
        ttk.Button(controls,text='Use automatic',command=self.clear_override).pack(side='left')
        self.details=tk.StringVar(); ttk.Label(preview,textvariable=self.details,wraplength=1080).pack(fill='x',pady=8)
        self.plotframe=ttk.Frame(preview); self.plotframe.pack(fill='both',expand=True)
        ttk.Label(self.root,textvariable=self.status,padding=10).pack(fill='x')
        self.root.after(100,self.drain_queue)
        import queue
        self.events=queue.Queue()

    def settings(self):
        s=Settings(**{k:float(v.get()) for k,v in self.variables.items()},common_y=self.common.get());s.validate();return s

    def error(self, exc):
        from tkinter import messagebox
        messagebox.showerror('Trypsin analysis',str(exc))

    def browse_input(self):
        from tkinter import filedialog
        p=filedialog.askopenfilename(filetypes=[('Excel workbook','*.xlsx')])
        if p:
            self.input.set(p);self.output.set(str(Path(p).with_name(Path(p).stem+'_v3_2.xlsx')))

    def browse_output(self):
        from tkinter import filedialog
        p=filedialog.asksaveasfilename(defaultextension='.xlsx',filetypes=[('Excel workbook','*.xlsx')])
        if p:self.output.set(p)

    def load(self):
        try:
            self.analyses=TrypsinWorkbookAnalyzer(self.input.get()).analyze();self.overrides={}
            self.loaded_source=self.input.get()
            self.pairs=[(a,o) for a in self.analyses for o in ('oligoT','oligoC','oligoA')]
            self.selector['values']=[f'{a.sheet_name} | {o}' for a,o in self.pairs]
            self.selector.current(0);self.refresh()
            self.status.set(f'Loaded {len(self.analyses)} worksheets and {len(self.pairs)} matched pairs.')
        except Exception as exc:self.error(exc)

    def selected_pair(self):
        i=self.selector.current()
        if i<0:raise ValueError('Load data first.')
        return self.pairs[i]

    def refresh(self,event=None):
        try:
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            a,o=self.selected_pair();s=self.settings();key=(a.sheet_name,o)
            r=analyze_pair(a,o,s,self.overrides.get(key))
            self.manual.set(str(self.overrides[key]) if key in self.overrides else '')
            self.details.set(f'{r.detection}. Window {r.start:g}–{r.end:g} min; '
                f'control slope {r.control[0]:.5g}, trypsin slope {r.trypsin[0]:.5g}; '
                f'dissolution rate {r.rate:.5g} AU/min. {r.review}')
            exporter=Exporter(self.analyses,s,self.loaded_source,self.overrides)
            fig=make_figure(r,exporter.ylim())
            if self.canvas:self.canvas.get_tk_widget().destroy()
            self.canvas=FigureCanvasTkAgg(fig,master=self.plotframe);self.canvas.draw()
            self.canvas.get_tk_widget().pack(fill='both',expand=True)
        except Exception as exc:self.error(exc)

    def apply_override(self):
        try:
            a,o=self.selected_pair();end=float(self.manual.get())
            analyze_pair(a,o,self.settings(),end)
            self.overrides[(a.sheet_name,o)]=end;self.refresh()
        except Exception as exc:self.error(exc)

    def clear_override(self):
        try:
            a,o=self.selected_pair();self.overrides.pop((a.sheet_name,o),None);self.refresh()
        except Exception as exc:self.error(exc)

    def export(self):
        try:
            if not self.analyses:raise ValueError('Load data first.')
            if self.input.get()!=self.loaded_source:raise ValueError('Input path changed. Load the new workbook first.')
            if not self.output.get().strip():raise ValueError('Choose an output workbook.')
            exporter=Exporter(self.analyses,self.settings(),self.loaded_source,dict(self.overrides))
            output=self.output.get();self.export_button.state(['disabled']);self.status.set('Creating workbook...')
            def worker():
                try:exporter.export(output);self.events.put(('ok',output))
                except Exception:self.events.put(('error',traceback.format_exc()))
            threading.Thread(target=worker,daemon=True).start()
        except Exception as exc:self.error(exc)

    def drain_queue(self):
        import queue
        try:
            while True:
                kind,value=self.events.get_nowait();self.export_button.state(['!disabled'])
                if kind=='ok':self.status.set('Saved: '+value)
                else:self.status.set('Export failed.');self.error(value)
        except queue.Empty:pass
        self.root.after(100,self.drain_queue)

    def run(self):self.root.mainloop()


def main():
    parser=argparse.ArgumentParser(description=SOFTWARE_VERSION)
    parser.add_argument('--input');parser.add_argument('--output')
    parser.add_argument('--start',type=float,default=5)
    parser.add_argument('--plateau-duration',type=float,default=10)
    parser.add_argument('--plateau-absolute',type=float,default=.005)
    parser.add_argument('--plateau-fraction',type=float,default=.02)
    parser.add_argument('--review-r2',type=float,default=.95)
    parser.add_argument('--independent-y',action='store_true')
    args=parser.parse_args()
    if args.input:
        if not args.output:parser.error('--output is required with --input')
        s=Settings(args.start,args.plateau_duration,args.plateau_absolute,args.plateau_fraction,args.review_r2,not args.independent_y)
        a=TrypsinWorkbookAnalyzer(args.input).analyze();e=Exporter(a,s,args.input);e.export(args.output)
        print(f'Saved {args.output}; {len(a)} worksheets, {len(e.results)} matched pairs.')
    else:TrypsinAssayApp().run()


if __name__=='__main__':main()
