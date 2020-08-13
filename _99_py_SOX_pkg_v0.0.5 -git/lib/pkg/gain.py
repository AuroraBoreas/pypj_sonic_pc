import os, sys, ctypes
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
from lib.util import pyver

def dll_gain(mic_rms, init_src_rms, center):
    if pyver.is_64bit():
        dll_path = os.path.join(os.path.dirname(__file__), "gain64.dll")
    else:
        dll_path = os.path.join(os.path.dirname(__file__), "gain32.dll")
    gainDll = ctypes.cdll.LoadLibrary(dll_path)
    # setup type of function arguments 
    gainDll.calc_gain.argtypes = [ctypes.c_float, ctypes.c_float, ctypes.c_float]
    # setup return type of function
    gainDll.calc_gain.restype = ctypes.c_float
    a = ctypes.c_float(mic_rms)
    b = ctypes.c_float(init_src_rms)
    c = ctypes.c_float(center)
    return round(gainDll.calc_gain(a, b, c), 2) 

def calc_gain(mic_rms, init_src_rms, center):
    return dll_gain(mic_rms, init_src_rms, center)