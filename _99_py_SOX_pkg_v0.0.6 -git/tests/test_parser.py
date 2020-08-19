import sys, os
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from lib.pkg.parser import read_sox_stats_and_get_diff_Pklev
import unittest

class ParserTestCase(unittest.TestCase):
    __base_dir = os.path.dirname(os.path.dirname(__file__))
    __src_file = os.path.join(__base_dir, r"data\sox_mic_src_stats")
    __diff = 5.67
    __src_Pklev = 25.34
    __lst = [26.79, 40.72, 25.34, 35.05]
    def test_diff(self):
        self.assertEqual(self.__diff, read_sox_stats_and_get_diff_Pklev(self.__src_file)[0])
    def test_srcPklev(self):
        self.assertEqual(self.__src_Pklev, read_sox_stats_and_get_diff_Pklev(self.__src_file)[1])
    def test_lst(self):
        self.assertEqual(self.__lst, read_sox_stats_and_get_diff_Pklev(self.__src_file)[2])

if __name__ == "__main__":
    unittest.main()