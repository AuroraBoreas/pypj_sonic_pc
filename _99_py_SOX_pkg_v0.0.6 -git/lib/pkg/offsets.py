"""
This module is to calculate "offset" of src_rms itself to compensate gain value.
Due to the nature of src_srm, it's bound with TV SW, it's stable.

Changelog
- v0.0.1 initial version

Author
@ZL, 20200819

"""
import sys, math, pathlib
sys.path.append(str(pathlib.Path(__file__).parent.parent.parent))
from lib import core
from lib.pkg import runbat
_Path = str

def get_offset(cali_bat: _Path) -> float:
    # get init_src_rms
    _gain         = 1
    _indx_src_rms = 3
    _tolerance    = 0.1
    # reset src_rms
    runbat.run_batch(batch_file=cali_bat)
    # start inital measurement. TODO: it's possible to get src_rms ONLY?
    _, _, micPkRMSsrcPkRMS = core.meas_and_get_result()
    init_src_rms = micPkRMSsrcPkRMS[_indx_src_rms]
    # change gain
    runbat.change_src_gain(gain_value=_gain)
    # get src_rms
    expc_src_rms = (init_src_rms - _gain)
    # actl_src_rms. TODO: it's possible to get src_rms ONLY?
    _, _, micPkRMSsrcPkRMS = core.meas_and_get_result()
    actl_src_rms = micPkRMSsrcPkRMS[_indx_src_rms]
    _isclose = math.isclose(expc_src_rms, actl_src_rms, abs_tol=_tolerance)
    if _isclose:
        offset = 0
    else:
        if abs(actl_src_rms) > abs(expc_src_rms):
            offset = abs(abs(expc_src_rms) - abs(actl_src_rms)) * -1
        else:
            offset = abs(abs(expc_src_rms) - abs(actl_src_rms))
    return offset