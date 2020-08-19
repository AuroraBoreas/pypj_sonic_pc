import os, sys, datetime
sys.path.append(os.path.dirname(__file__))
from lib import core

def single_meas():
    recorder = "sox_stats_log.txt"
    header   = "Diff, src_25_Pk, [mic_25_Pk, mic_25_RMS, src_25_Pk, src_25_RMS]"
    result   = core.meas_and_get_result()
    timer    = datetime.datetime.now

    with open(recorder, 'a') as f:
        f.writelines("{0:%Y-%m-%d %H:%M:%S} {1}\n".format(timer(), header))
        f.writelines("{0:%Y-%m-%d %H:%M:%S} {1}\n".format(timer(), result))
    return

if __name__ == "__main__":
    single_meas()