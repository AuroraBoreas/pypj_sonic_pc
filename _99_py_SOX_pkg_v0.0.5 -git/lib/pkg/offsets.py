def calc_offset(mic_rms_before: float, mic_rms_after: float, act_src_rms: float, exp_src_rms: float) -> float:
    diff = abs(act_src_rms) - abs(exp_src_rms)
    # assumes *change_src_gain* works on src_rms only
    # speculate TV src_rms *change_src_gain* causes linear offset
    # so, it means offset has no direct impact from mic_rms, src_rms, and input_src_gain
    # unification
    mic_rms_variation = abs(abs(mic_rms_before) - abs(mic_rms_after))
    bound = 0.1
    # if mic_rms variation is less than 0.1, do compensation .. 
    if mic_rms_variation <= bound:
        if diff > 0:
            offset = abs(diff)
        elif diff < 0:
            offset = -1 * abs(diff)
        else:
            offset = 0
    # if mic_rms varies drastically, NO compensation ..
    else:
        offset = 0
    return round(offset, 2)