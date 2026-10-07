"""Run complete-program synthetic examples after installing requirements.

Requires a desktop display for the hidden Tk droplet smoke test. Exit 2 means
dependencies are missing; exit 1 means a check failed. Never a manuscript validation.
"""
from __future__ import annotations
import importlib.util
import json
import math
import os
from pathlib import Path
import shutil
import sys
import traceback
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'outputs' / 'example_validation'

def load(name, relative):
    spec = importlib.util.spec_from_file_location(name, ROOT / relative)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module

def run_frap():
    import numpy as np
    import openpyxl
    f = load('pnh_frap_example','scripts/frap/frap_gui_app_v5_with_frame_finder.py')
    folder = OUT / 'frap'
    folder.mkdir(exist_ok=True)
    source = folder / 'example.csv'
    shutil.copyfile(ROOT/'examples/synthetic/frap/example.csv',source)
    f.PLOTS_DIR = folder / 'plots'
    f.PLOTS_DIR.mkdir(exist_ok=True)
    target = folder / 'summary.xlsx'
    sample = f.SampleInput('synthetic_frap',folder,[source],target)
    f.run_analysis_for_sample(sample,'auto','fixed',90.0)
    book = openpyxl.load_workbook(target,data_only=True)
    sheet = book['half_life']
    headers = [cell.value for cell in sheet[1]]
    half = sheet.cell(2,headers.index('Channel1_half_life_s')+1).value
    if not math.isclose(half,20*math.log(2),rel_tol=1e-5):
        raise AssertionError(f'Unexpected FRAP half-time {half}')
    book.close()
    # A!=B differentiates half-plateau crossing from half-amplitude recovery.
    x = np.arange(121,dtype=float)
    result = f._fit_half_life_once(x,100-80*np.exp(-x/20),90)
    expected = math.log(.5*100/80)/(-1/20)
    if not math.isclose(result['half_life_s'],expected,rel_tol=1e-5):
        raise AssertionError('Half-plateau definition mismatch')
    return {'workbook':str(target.relative_to(ROOT)), 'half_time_s':half,
            'offset_curve_half_time_s':result['half_life_s']}

def run_trypsin():
    import openpyxl
    t = load('pnh_trypsin_example','scripts/trypsin/trypsin_assay_gui_v4.py')
    folder = OUT / 'trypsin'
    folder.mkdir(exist_ok=True)
    source = folder / 'example.xlsx'
    shutil.copyfile(ROOT/'examples/synthetic/trypsin/example.xlsx',source)
    analyses = t.TrypsinWorkbookAnalyzer(source).analyze()
    exporter = t.Exporter(analyses,t.Settings(),source)
    target = folder / 'summary.xlsx'
    exporter.export(target)
    for result in exporter.results:
        expected = .004 if result.analysis.iteration == 1 else .006
        if not math.isclose(result.rate,expected,abs_tol=1e-10):
            raise AssertionError('Trypsin rate mismatch')
    book = openpyxl.load_workbook(target,data_only=True)
    for row in book['Summary'].iter_rows(min_row=2,max_row=4,values_only=True):
        if not math.isclose(row[3],.005,abs_tol=1e-10):
            raise AssertionError('Exported mean mismatch')
        if not math.isclose(row[4],math.sqrt(2)*.001,abs_tol=1e-10):
            raise AssertionError('Exported sample SD mismatch')
    book.close()
    return {'workbook':str(target.relative_to(ROOT)), 'matched_pairs':len(exporter.results)}

def run_droplets():
    import tkinter as tk
    import openpyxl
    d = load('pnh_droplet_example','scripts/droplet_counting/droplet_counter_gui_v6_1_fast_preview.py')
    folder = OUT / 'droplets'
    folder.mkdir(exist_ok=True)
    # Isolate test preset location; do not touch the user's actual saved presets.
    class TestApp(d.DropletCounterApp):
        @property
        def preset_file(self):
            return folder/'test_presets.json'
    root = tk.Tk()
    root.withdraw()
    def fail_dialog(*args,**kwargs):
        raise AssertionError('Application export error: '+str(args))
    try:
        app = TestApp(root)
        frames = []
        for path in sorted((ROOT/'examples/synthetic/droplets').glob('*.png')):
            frames.extend(d.load_image_file(path))
        app.frames = frames
        counts = [app.analyze_frame(f.image,i,f.name).count for i,f in enumerate(frames)]
        (folder/'effective_parameters.json').write_text(json.dumps(app.get_params_values(),indent=2),encoding='utf-8')
        if counts != [0,4]:
            raise AssertionError(f'Detector challenge failed: expected [0,4], observed {counts}. Inspect overlays; do not retune silently.')
        target = folder/'summary.xlsx'
        with patch.object(d.filedialog,'asksaveasfilename',return_value=str(target)), \
             patch.object(d.messagebox,'showinfo'), patch.object(d.messagebox,'showerror',side_effect=fail_dialog):
            app.export_excel()
        book = openpyxl.load_workbook(target,data_only=True)
        sheet = book['per_frame_summary']
        headers = [c.value for c in sheet[1]]
        exported = [sheet.cell(i,headers.index('droplet_count')+1).value for i in (2,3)]
        book.close()
        if exported != counts:
            raise AssertionError('Exported counts differ')
        return {'workbook':str(target.relative_to(ROOT)), 'counts':counts}
    finally:
        root.destroy()

def main():
    OUT.mkdir(parents=True,exist_ok=True)
    modules = ['numpy','pandas','scipy','matplotlib','openpyxl','xlsxwriter','PIL','tifffile','skimage','tkinter']
    missing = [name for name in modules if importlib.util.find_spec(name) is None]
    report = {'synthetic_only':True,'python':sys.version,'missing_dependencies':missing,'checks':{}}
    if missing:
        report['status'] = 'BLOCKED'
        print('Full-program examples BLOCKED; missing:',', '.join(missing))
        exit_code = 2
    else:
        os.environ['MPLBACKEND'] = 'Agg'
        for name, action in [('frap',run_frap),('trypsin',run_trypsin),('droplets',run_droplets)]:
            try:
                report['checks'][name] = {'status':'PASS','result':action()}
            except Exception:
                report['checks'][name] = {'status':'FAIL','error':traceback.format_exc()}
        exit_code = 0 if all(v['status']=='PASS' for v in report['checks'].values()) else 1
        report['status'] = 'PASS' if exit_code == 0 else 'FAIL'
        print(json.dumps(report,indent=2))
    (OUT/'report.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
    return exit_code

if __name__ == '__main__':
    sys.exit(main())
