import math
from pathlib import Path
import unittest
import numpy as np
from core_loader import ROOT, frap, trypsin, load_selected

class ComputationalChecks(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.t = trypsin()
        cls.f = frap()

    def test_trypsin_workbook_grouping_units_and_rates(self):
        analyses = self.t.TrypsinWorkbookAnalyzer(ROOT/'examples/synthetic/trypsin/example.xlsx').analyze()
        self.assertEqual(len(analyses),2)
        rates = []
        for a,expected in zip(analyses,[.004,.006]):
            self.assertEqual(a.duration_minutes,40)
            self.assertEqual(a.raw_values.shape,(41,18))
            for dna in ['oligoT','oligoC','oligoA']:
                r = self.t.analyze_pair(a,dna,self.t.Settings())
                self.assertAlmostEqual(r.rate,expected,places=12)
                self.assertEqual(r.start,5)
                self.assertIn('plateau',r.detection.lower())
            rates.append(r.rate)
        self.assertAlmostEqual(np.mean(rates),.005,places=12)
        self.assertAlmostEqual(np.std(rates,ddof=1),math.sqrt(2)*.001,places=12)

    def test_trypsin_linear_slope_and_missing_points(self):
        t = np.arange(10,dtype=float)
        y = .8-.004*t
        y[3] = np.nan
        m,b,r2,n = self.t.linear_fit(t,y)
        self.assertAlmostEqual(m,-.004,places=12)
        self.assertAlmostEqual(b,.8,places=12)
        self.assertEqual(n,9)
        self.assertAlmostEqual(r2,1)

    def test_trypsin_manual_endpoint(self):
        a = self.t.TrypsinWorkbookAnalyzer(ROOT/'examples/synthetic/trypsin/example.xlsx').analyze()[0]
        r = self.t.analyze_pair(a,'oligoT',self.t.Settings(),override=15)
        self.assertEqual(r.end,15)
        self.assertEqual(r.detection,'Manual fit end override')
        self.assertAlmostEqual(r.rate,.004,places=12)

    def test_trypsin_no_plateau_fits_to_end(self):
        t = np.arange(41,dtype=float)
        end,tol,status = self.t.detect_plateau(t,.8-.005*t,self.t.Settings())
        self.assertEqual(end,40)
        self.assertIn('fitted to end',status)

    def test_trypsin_negative_rate_retained(self):
        a = self.t.TrypsinWorkbookAnalyzer(ROOT/'examples/synthetic/trypsin/example.xlsx').analyze()[0]
        a.group('oligoT','Control').mean = .8-.008*a.time_minutes
        r = self.t.analyze_pair(a,'oligoT',self.t.Settings(),override=20)
        self.assertAlmostEqual(r.rate,-.003,places=12)
        self.assertIn('Negative dissolution rate retained',r.review)

    def test_frap_csv_correction_and_normalization(self):
        exp = self.f.parse_frap_file(ROOT/'examples/synthetic/frap/example.csv')
        self.assertEqual(exp.bleach_time_seconds,5)
        processed = self.f.compute_processed_trace(exp,'Channel1')
        norm = self.f.normalize_trace(exp.time_seconds,5,processed)
        expected = 100*(1-np.exp(-norm.time_from_bleach_s.to_numpy()/20))
        np.testing.assert_allclose(norm.normalized_intensity,expected,rtol=1e-12,atol=1e-12)

    def test_frap_zero_reference_denominator(self):
        exp = self.f.parse_frap_file(ROOT/'examples/synthetic/frap/example.csv')
        exp.channels['Channel1']['R2'] = exp.channels['Channel1']['R3'].copy()
        processed = self.f.compute_processed_trace(exp,'Channel1')
        self.assertTrue(np.isnan(processed).all())

    def test_frap_background_only_fallback(self):
        exp = self.f.parse_frap_file(ROOT/'examples/synthetic/frap/example.csv')
        del exp.channels['Channel1']['R2']
        processed = self.f.compute_processed_trace(exp,'Channel1')
        np.testing.assert_allclose(processed,exp.channels['Channel1']['R1']-10)

    def test_frap_uses_extrema_not_first_post_point(self):
        t = np.array([0.,1.,2.,3.,4.])
        y = np.array([80.,100.,30.,20.,60.])
        result = self.f.normalize_trace(t,2,y)
        np.testing.assert_allclose(result.normalized_intensity,[12.5,0,50])

    def test_droplet_filename_temperature_only(self):
        mod = load_selected('droplets',['extract_temperature_c_from_name'],[])
        self.assertEqual(mod.extract_temperature_c_from_name('25_76C_2W-oT_cropped'),76)
        self.assertIsNone(mod.extract_temperature_c_from_name('no_temperature'))

if __name__ == '__main__':
    unittest.main()
