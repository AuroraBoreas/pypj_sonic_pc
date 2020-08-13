import unittest
import sys, os
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from lib.pkg import offsets

class GainTestCase(unittest.TestCase):
    def test_calc_gain1(self):
        mic_rms_before, mic_rms_after = 44.55, 44.50
        act_src_rms, exp_src_rms = 36.66, 36.55
        result = offsets.calc_offset(mic_rms_before, mic_rms_after, act_src_rms, exp_src_rms)
        self.assertEqual(0.11, result)
    def test_calc_gain2(self):
        mic_rms_before, mic_rms_after = -44.55, 44.50
        act_src_rms, exp_src_rms = 36.66, 36.55
        result = offsets.calc_offset(mic_rms_before, mic_rms_after, act_src_rms, exp_src_rms)
        self.assertEqual(0.11, result)
    def test_calc_gain3(self):
        mic_rms_before, mic_rms_after = -44.55, -44.50
        act_src_rms, exp_src_rms = 36.66, 36.55
        result = offsets.calc_offset(mic_rms_before, mic_rms_after, act_src_rms, exp_src_rms)
        self.assertEqual(0.11, result)
    def test_calc_gain4(self):
        mic_rms_before, mic_rms_after = 44.55, -44.50
        act_src_rms, exp_src_rms = 36.66, 36.55
        result = offsets.calc_offset(mic_rms_before, mic_rms_after, act_src_rms, exp_src_rms)
        self.assertEqual(0.11, result)
    def test_calc_gain5(self):
        mic_rms_before, mic_rms_after = 44.55, -44.50
        act_src_rms, exp_src_rms = -36.66, 36.55
        result = offsets.calc_offset(mic_rms_before, mic_rms_after, act_src_rms, exp_src_rms)
        self.assertEqual(0.11, result)
    def test_calc_gain6(self):
        mic_rms_before, mic_rms_after = 44.55, -44.50
        act_src_rms, exp_src_rms = 36.66, -36.55
        result = offsets.calc_offset(mic_rms_before, mic_rms_after, act_src_rms, exp_src_rms)
        self.assertEqual(0.11, result)
    def test_calc_gain7(self):
        mic_rms_before, mic_rms_after = 44.55, 44.80
        act_src_rms, exp_src_rms = 36.66, -36.55
        result = offsets.calc_offset(mic_rms_before, mic_rms_after, act_src_rms, exp_src_rms)
        self.assertEqual(0.00, result)

if __name__ == "__main__":
    unittest.main()