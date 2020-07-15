
def calc_gain(mic_rms, init_src_rms, center):
    diff = round(abs(mic_rms - init_src_rms), 2)
    gain = round(abs(center - diff), 2)
    cond1 = abs(mic_rms) - abs(init_src_rms)
    cond2 = diff - center

    if (cond1 > 0 and cond2 > 0) or (cond1 < 0 and cond2 < 0):
        return -1 * gain

    if (cond1 > 0 and cond2 < 0) or (cond1 < 0 and cond2 > 0):
        return gain