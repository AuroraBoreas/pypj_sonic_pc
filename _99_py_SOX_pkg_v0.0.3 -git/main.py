"""A module measures remote field sound automatically

It has the following functionalities
- step1, step2, step3 to measure and analyze

Changelog
- v0.0.1, initial fork
- v0.0.2, refactor and display runtime of this program
- v0.0.3, format output
- v0.0.4, refactor and include a simple GUI

About
- branch : forked by @ZL, 20200630
"""

import time, datetime, statistics, os, sys
sys.path.append(os.path.dirname(__file__))
from lib import core
from lib.pkg import runbat, gain
import tkinter
from tkinter import messagebox

class RemoteFieldSoundSearcher:
    """A class search gain value between Reference signal and Mic signal to SONY TV remote field sound
    """
    # formatter and timer
    __timer      = datetime.datetime.now
    __info_fmt   = "{0:<12}: {1:%Y-%m-%d %H:%M:%S},Diff={2:<.2f},mic_RMS={3:<.2f},src_RMS={4:<.2f},Gain_val={5:<.2f}"
    __header     = "gain_val, [mic_pk, mic_rms, src_pk, src_rms]"
    __result_log = "m5_log.txt"
    # batch file for calibO
    __base_dir1       = r"C:\Users\ssv\Desktop\m5\bat"
    __bat_file_calib0 = os.path.join(__base_dir1, "20200702_calib0.bat")
    # magic numbers from SHES Design
    __center    = 6.0
    __lbound    = __center - 0.1
    __ubound    = __center + 0.1
    __max_val   = 144
    __init_gain = 0
    # initialize containers to store results
    __result_list  = []
    __mic_rms_list = []

    def __init__(self, max_search=3):
        self._maxsearch = max_search
    def _save(self):
        """write search result into local file
        """
        with open(self.__result_log, "a") as f:
            f.writelines("{0:%Y-%m-%d %H:%M:%S} {1}\n".format(self.__timer(), self.__header))
            f.writelines("{0:%Y-%m-%d %H:%M:%S} {1}\n".format(self.__timer(), self.__result_list))
        return
    def mease_and_search(self):
        """measure and search the closest gain value against *__center*
        """
        cnt = 0
        start_time = time.time()
        # reset Reference
        runbat.run_batch(batch_file=self.__bat_file_calib0)
        # start inital measurement
        srcdiff, srcpeak, micPkRMSsrcPkRMS = core.meas_and_get_result()
        init_src_rms = micPkRMSsrcPkRMS[3]
        # display result of initial measurement
        print(self.__info_fmt.format("init meas", self.__timer(), srcdiff, micPkRMSsrcPkRMS[1], micPkRMSsrcPkRMS[3], self.__init_gain))
        # if srcdiff < lbound or srcdiff > ubound ...
        if self.__lbound <= srcdiff <= self.__ubound:
            self.__result_list.append((self.__init_gain, micPkRMSsrcPkRMS))
        else: 
            while not (self.__lbound <= srcdiff <= self.__ubound):
                input_src_gain = gain.calc_gain(mic_rms=micPkRMSsrcPkRMS[1], init_src_rms=init_src_rms, center=self.__center)
                self.__mic_rms_list.append(micPkRMSsrcPkRMS[1])
                # if src Pk lev dB == 0 ...
                if srcpeak == 0:
                    pass
                # if input src gain vaue > max_val ...    
                if input_src_gain > self.__max_val:
                    pass
                if cnt >= self._maxsearch:
                    # after searching N times, but no result, calculate harmonic mean of mic_25_rms values
                    hm_mic25rms    = round(statistics.harmonic_mean(self.__mic_rms_list), 2)
                    input_src_gain = gain.calc_gain(mic_rms=hm_mic25rms, init_src_rms=init_src_rms, center=self.__center)
                    runbat.change_src_gain(gain_value=input_src_gain) # change gain value
                    result = input_src_gain
                    _, _, micPkRMSsrcPkRMS = core.meas_and_get_result()
                    micPkRMSsrcPkRMS[1] = hm_mic25rms
                    break
                # display result before change gain
                print(self.__info_fmt.format("before gain", self.__timer(), srcdiff, micPkRMSsrcPkRMS[1], micPkRMSsrcPkRMS[3], input_src_gain))
                # change gain value
                runbat.change_src_gain(gain_value=input_src_gain) 
                result = input_src_gain
                # measure after changed gain value
                srcdiff, srcpeak, micPkRMSsrcPkRMS = core.meas_and_get_result() 
                # display result after change gain
                print(self.__info_fmt.format("after gain", self.__timer(), srcdiff, micPkRMSsrcPkRMS[1], micPkRMSsrcPkRMS[3], input_src_gain))
                cnt += 1
            self.__result_list.append((result, micPkRMSsrcPkRMS))
        end_time = time.time()
        print("Finished. Runtime is {0:.2f}s".format(end_time - start_time))
        print("Result:", self.__result_list)
        self._save()
        return

def main():
    searcher = RemoteFieldSoundSearcher()
    searcher.mease_and_search()

if __name__ == "__main__":
    root = tkinter.Tk()
    root.withdraw()
    tv_sound_volumes = [1, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 70, 80, 90, 100]
    for volume in tv_sound_volumes:
        user_answer = messagebox.askquestion(title="TV Sound Volume Searcher", 
                                            message="Start to search volume={0}".format(volume),
                                            icon="warning", 
                                            type="okcancel")
        if user_answer == "ok":
            main()
        else:
            pass
    root.destroy()