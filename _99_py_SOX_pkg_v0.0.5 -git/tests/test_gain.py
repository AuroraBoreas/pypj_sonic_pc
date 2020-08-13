import unittest
import sys, os
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from lib.pkg import gain

class GainTestCase(unittest.TestCase):
    def test_calc_gain1(self):
        center = 6.0 
        mic_rms, init_src_rms = -37.41, -30.85
        result = gain.calc_gain(mic_rms, init_src_rms, center)
        self.assertEqual(-0.56, result)
    def test_calc_gain2(self):
        center = 6.0 
        mic_rms, init_src_rms = -40.59, -35.06
        result = gain.calc_gain(mic_rms, init_src_rms, center)
        self.assertEqual(0.47, result)
    def test_calc_gain3(self):
        center = 6.0 
        mic_rms, init_src_rms = -45.05, -36.66
        result = gain.calc_gain(mic_rms, init_src_rms, center)
        self.assertEqual(-2.39, result)
    def test_calc_gain4(self):
        center = 6.0 
        mic_rms, init_src_rms = -37.43, -30.86
        result = gain.calc_gain(mic_rms, init_src_rms, center)
        self.assertEqual(-0.57, result)
    def test_calc_gain5(self):
        center = 6.0 
        mic_rms, init_src_rms = -33.47, -27.65
        result = gain.calc_gain(mic_rms, init_src_rms, center)
        self.assertEqual(0.18, result)
    def test_calc_gain6(self):
        center = 6.0 
        mic_rms, init_src_rms = 14.22, -19.91
        result = gain.calc_gain(mic_rms, init_src_rms, center)
        self.assertEqual(11.69, result)
    def test_calc_gain7(self):
        center = 6.0 
        mic_rms, init_src_rms = -13.48, 14.91
        result = gain.calc_gain(mic_rms, init_src_rms, center)
        self.assertEqual(7.43, result)
    def test_calc_gain8(self):
        center = 6.0 
        mic_rms, init_src_rms = -13.25, -14.91
        result = gain.calc_gain(mic_rms, init_src_rms, center)
        self.assertEqual(7.66, result)
    def test_calc_gain9(self):
        center = 6.0 
        mic_rms, init_src_rms = 21.43, 49.94
        result = gain.calc_gain(mic_rms, init_src_rms, center)
        self.assertEqual(34.51, result)
    def test_calc_gain10(self):
        center = 6.0 
        mic_rms, init_src_rms = -21.43, 21.43
        result = gain.calc_gain(mic_rms, init_src_rms, center)
        self.assertEqual(6.0, result)
    def test_calc_gain11(self):
        center = 6.0 
        mic_rms, init_src_rms = 21.43, -21.43
        result = gain.calc_gain(mic_rms, init_src_rms, center)
        self.assertEqual(6.0, result)
    def test_calc_gain12(self):
        center = 6.0 
        mic_rms, init_src_rms = 21.43, 21.43
        result = gain.calc_gain(mic_rms, init_src_rms, center)
        self.assertEqual(6.0, result)

if __name__ == "__main__":
    unittest.main()