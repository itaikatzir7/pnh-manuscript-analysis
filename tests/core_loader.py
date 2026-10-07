"""Load selected original definitions for limited dependency-independent tests.

This deliberately does not exercise imports, GUI initialization, plotting or SciPy fits.
No functions are rewritten or mocked. Full-program tests are in tools/run_examples.py.
"""
import ast
import dataclasses
import math
import os
from pathlib import Path
import re
import sys
import types
from typing import Any, Dict, Iterable, List, Optional, Tuple
import zipfile
from xml.etree import ElementTree as ET
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]

def load_selected(kind, names, constants):
    paths = {'frap':'scripts/frap/frap_gui_app_v5_with_frame_finder.py',
             'trypsin':'scripts/trypsin/trypsin_assay_gui_v4.py',
             'droplets':'scripts/droplet_counting/droplet_counter_gui_v6_1_fast_preview.py'}
    path = ROOT / paths[kind]
    tree = ast.parse(path.read_text(encoding='utf-8-sig'))
    selected = []
    for node in tree.body:
        if isinstance(node,(ast.FunctionDef,ast.ClassDef)) and node.name in names:
            selected.append(node)
        elif isinstance(node,ast.Assign) and any(isinstance(t,ast.Name) and t.id in constants for t in node.targets):
            selected.append(node)
    present = {n.name for n in selected if isinstance(n,(ast.FunctionDef,ast.ClassDef))}
    if present != set(names):
        raise ValueError('Unresolved selected definitions: '+str(set(names)-present))
    module = types.ModuleType('selected_'+kind)
    sys.modules[module.__name__] = module
    module.__dict__.update(np=np,pd=pd,math=math,os=os,Path=Path,re=re,zipfile=zipfile,ET=ET,
                           dataclass=dataclasses.dataclass,field=dataclasses.field,
                           Any=Any,Dict=Dict,Iterable=Iterable,List=List,Optional=Optional,Tuple=Tuple)
    exec(compile(ast.Module(body=selected,type_ignores=[]),str(path),'exec'),module.__dict__)
    return module

def trypsin():
    return load_selected('trypsin',[
        '_local_name','_excel_col_to_index','_excel_index_to_col','_safe_float',
        '_parse_time_string_to_minutes','_time_raw_scalar','_is_all_zero_or_blank',
        '_nanstd_triplicates','_baseline','_parse_system_iteration','_extract_wavelength',
        'XlsxRawReader','GroupData','SheetData','TrypsinWorkbookAnalyzer','Settings',
        'linear_fit','detect_plateau','PairResult','analyze_pair'],
        ['GROUP_SPECS','MAIN_NS','REL_NS','PKG_REL_NS'])

def frap():
    return load_selected('frap',[
        'FRAPExperiment','unit_to_seconds_multiplier','read_input_file',
        'pick_relevant_intensity_columns','parse_channel_and_roi','parse_frap_file',
        'compute_processed_trace','normalize_trace'],['CHANNEL_PATTERN','MIN_DENOMINATOR_ABS'])
