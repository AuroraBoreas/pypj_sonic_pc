import os, ctypes

def dll_gain(mic_rms, init_src_rms, center):
    dll_path = os.path.join(os.path.dirname(__file__), "gain.dll")
    gainDll = ctypes.cdll.LoadLibrary(dll_path)
    # setup type of function arguments 
    gainDll.calc_gain.argtypes = [ctypes.c_float, ctypes.c_float, ctypes.c_float]
    # setup return type of function
    gainDll.calc_gain.restype = ctypes.c_float
    a = ctypes.c_float(mic_rms)
    b = ctypes.c_float(init_src_rms)
    c = ctypes.c_float(center)
    return round(gainDll.calc_gain(a, b, c),2)

def py_gain(mic_rms, init_src_rms, center):
    diff = round(abs(mic_rms - init_src_rms), 2)
    gain = round(abs(center - diff), 2)
    cond1 = abs(mic_rms) - abs(init_src_rms)
    cond2 = diff - center
    
    if cond1 > 0:
        return -1 * gain if cond2 > 0 else gain
    if cond1 < 0:
        return (-1 * gain + center * 2) if cond2 < 0 else (gain + center * 2) 

def calc_gain(mic_rms, init_src_rms, center):
    try:
        return dll_gain(mic_rms, init_src_rms, center)
    except:
        return py_gain(mic_rms, init_src_rms, center)