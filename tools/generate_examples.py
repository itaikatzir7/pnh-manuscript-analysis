"""Generate deterministic, synthetic examples. No manuscript data."""
import csv
import json
import math
from pathlib import Path
from PIL import Image, ImageDraw
import xlsxwriter

ROOT = Path(__file__).resolve().parents[1]

def main():
    target = ROOT / 'examples' / 'synthetic'
    for name in ('frap','trypsin','droplets'):
        (target / name).mkdir(parents=True, exist_ok=True)
    with (target / 'frap' / 'example.csv').open('w', newline='', encoding='utf-8') as handle:
        writer = csv.writer(handle)
        writer.writerow(['Time','Marker','Channel1_R1_IntensityMean','Channel1_R2_IntensityMean','Channel1_R3_IntensityMean'])
        writer.writerow(['s','','','',''])
        for time in range(126):
            intensity = 110 if time < 5 else 30 + 80 * (1-math.exp(-(time-5)/20))
            writer.writerow([time,'bleach' if time == 5 else '',intensity,210,10])
    with xlsxwriter.Workbook(target / 'trypsin' / 'example.xlsx') as book:
        for iteration, slope in enumerate((.005,.007),1):
            sheet = book.add_worksheet(f'synthetic_{iteration}')
            sheet.write_row(0,0,['Time','Temperature'] + [f'A{i}' for i in range(1,19)])
            for time in range(41):
                values = []
                for _ in range(3):
                    values.extend(.8 - .001*time + offset for offset in (-.01,0,.01))
                    values.extend(.8 - slope*min(time,20) + offset for offset in (-.01,0,.01))
                sheet.write_row(time+1,0,[f'00:{time:02d}:00',25,*values])
    blank = Image.new('L',(128,128),200)
    blank.save(target / 'droplets' / '01_90C_synthetic_blank.png')
    disks = blank.copy()
    draw = ImageDraw.Draw(disks)
    centers = [(32,32),(96,32),(32,96),(96,96)]
    for x,y in centers:
        draw.ellipse((x-6,y-6,x+6,y+6),fill=40)
    disks.save(target / 'droplets' / '02_25C_synthetic_disks.png')
    expected = {'synthetic':True,'frap':{'bleach_time_s':5,'tau_s':20,'half_time_s':20*math.log(2)},
                'trypsin':{'iteration_rates_AU_per_min':[.004,.006],'mean_AU_per_min':.005,'sample_sd_AU_per_min':math.sqrt(2)*.001},
                'droplets':{'blank_count':0,'disk_count':4,'disk_centers_xy':centers,'radius_px':6,
                            'note':'Ideal object truth, not an asserted detector result. Validate actual detections.'}}
    (target / 'expected.json').write_text(json.dumps(expected,indent=2)+'\n',encoding='utf-8')
    print('Synthetic examples generated. None are manuscript measurements.')

if __name__ == '__main__':
    main()
