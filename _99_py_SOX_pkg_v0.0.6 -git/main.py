"""A module measures remote field sound automatically

It has the following functionalities
- step1, step2, step3 to measure and analyze

Changelog
- v0.0.1, initial fork
- v0.0.2, refactor and display runtime of this program
- v0.0.3, format output
- v0.0.4, refactor and include a simple GUI
- v0.0.5, compensate src_rms offset to gain value
- v0.0.6, improve algorithm of src_rms offset calculation

About
- branch : forked by @ZL, 20200630
"""

import time, datetime, statistics, os, sys
sys.path.append(os.path.dirname(__file__))
from lib import core
from lib.pkg import runbat, gain, offsets
import tkinter
from tkinter import messagebox
from scipy import stats

class RemoteFieldSoundSearcher:
    """A class search gain value between Reference signal and Mic signal to SONY TV remote field sound
    """
    # formatter and timer
    __timer      = datetime.datetime.now
    __info_fmt   = "{0:<12}: {1:%H:%M:%S},Diff={2:<.2f},mic_RMS={3:<.2f},src_RMS={4:<.2f},Gain_val={5:<.2f},exp_src_RMS={6:<.2f}"
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
    __indx_mic_rms = 1
    __indx_src_rms = 3
    __tm_proport   = .2
    __prec         = 2

    def __init__(self, max_search=5):
        self._maxsearch = max_search
    def __save(self):
        """write search result into local file
        """
        with open(self.__result_log, "a") as f:
            f.writelines("{0:%Y-%m-%d %H:%M:%S} {1}\n".format(self.__timer(), self.__header))
            f.writelines("{0:%Y-%m-%d %H:%M:%S} {1}\n".format(self.__timer(), self.__result_list))
        return
    def __saveSearchLog(self, s):
        with open(self.__result_log, "a") as f:
            f.writelines(s + "\n")
        return
    def start(self):
        """measure and search the closest gain value against *__center*
        """
        start_time = time.time()
        # get offset
        src_rms_offset = offsets.get_offset(cali_bat=self.__bat_file_calib0)
        # reset Reference
        runbat.run_batch(batch_file=self.__bat_file_calib0)
        # start inital measurement
        srcdiff, srcpeak, micPkRMSsrcPkRMS = core.meas_and_get_result()
        init_src_rms = micPkRMSsrcPkRMS[self.__indx_src_rms]
        # display result of initial measurement
        s = self.__info_fmt.format("init meas", self.__timer(), srcdiff, 
                                    micPkRMSsrcPkRMS[self.__indx_mic_rms], 
                                    micPkRMSsrcPkRMS[self.__indx_src_rms], 
                                    self.__init_gain,
                                    micPkRMSsrcPkRMS[self.__indx_src_rms])
        print(s)
        self.__saveSearchLog(s)
        # if srcdiff < lbound or srcdiff > ubound ...
        if self.__lbound <= srcdiff <= self.__ubound:
            self.__result_list.append((self.__init_gain, micPkRMSsrcPkRMS))
        else: 
            cnt = 0
            while not self.__lbound <= srcdiff <= self.__ubound:
                input_src_gain = gain.calc_gain(mic_rms=micPkRMSsrcPkRMS[self.__indx_mic_rms], init_src_rms=init_src_rms, center=self.__center) + src_rms_offset
                self.__mic_rms_list.append(micPkRMSsrcPkRMS[self.__indx_mic_rms])
                # if src Pk lev dB == 0 ...
                if srcpeak == 0:
                    pass
                # if input src gain vaue > max_val ...    
                if input_src_gain > self.__max_val:
                    pass
                if cnt >= self._maxsearch:
                    # after searching N times, but no result, calculate trimmed mean of mic_25_rms values
                    tm_mic25rms    = round(stats.trim_mean(self.__mic_rms_list, self.__tm_proport), self.__prec)
                    input_src_gain = gain.calc_gain(mic_rms=tm_mic25rms, init_src_rms=init_src_rms, center=self.__center) + src_rms_offset
                    runbat.change_src_gain(gain_value=input_src_gain) # change gain value
                    result = input_src_gain
                    _, _, micPkRMSsrcPkRMS = core.meas_and_get_result()
                    micPkRMSsrcPkRMS[self.__indx_mic_rms] = tm_mic25rms
                    break
                # display result before change gain
                exp_src_RMS = (init_src_rms + input_src_gain) if input_src_gain < 0 else (init_src_rms - input_src_gain)
                s = self.__info_fmt.format("before gain", self.__timer(), srcdiff, 
                                            micPkRMSsrcPkRMS[self.__indx_mic_rms], 
                                            micPkRMSsrcPkRMS[self.__indx_src_rms], 
                                            input_src_gain,
                                            exp_src_RMS)
                print(s)
                self.__saveSearchLog(s)
                # change gain value
                runbat.change_src_gain(gain_value=input_src_gain) 
                result = input_src_gain
                # measure after changed gain value
                srcdiff, srcpeak, micPkRMSsrcPkRMS = core.meas_and_get_result()
                # cond2 = (micPkRMSsrcPkRMS[self.__indx_mic_rms] - micPkRMSsrcPkRMS[self.__indx_src_rms]) > 0
                # display result after change gain
                exp_src_RMS = (init_src_rms + input_src_gain) if input_src_gain < 0 else (init_src_rms - input_src_gain)
                s = self.__info_fmt.format("after gain", self.__timer(), srcdiff, 
                                            micPkRMSsrcPkRMS[self.__indx_mic_rms], 
                                            micPkRMSsrcPkRMS[self.__indx_src_rms], 
                                            input_src_gain,
                                            exp_src_RMS)
                print(s)
                self.__saveSearchLog(s)
                cnt += 1
            self.__result_list.append((result, micPkRMSsrcPkRMS))
        end_time = time.time()
        print("Finished. Runtime is {0:.2f}s".format(end_time - start_time))
        print("Result:", self.__result_list)
        self.__save()
        return

def main():
    searcher = RemoteFieldSoundSearcher()
    searcher.start()

if __name__ == "__main__":
    root = tkinter.Tk()
    root.withdraw()
    # tv_sound_volumes = [1, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 70, 80, 90, 100]
    # for volume in tv_sound_volumes:
    user_answer = messagebox.askquestion(title="TV Sound Volume Searcher", 
                                        message="Start to search",
                                        icon="warning", 
                                        type="okcancel")
    if user_answer == "ok":
        try:
            main()
        except FileNotFoundError:
            messagebox.showinfo(title="TV Remote Field Audio Auto", message="PD Jig PC Only", icon="warning")
            sys.exit("batch files not found")
        except IndexError:
            messagebox.showinfo(title="TV Remote Field Audio Auto", message="SoX Failed", icon="warning")
            sys.exit("SoX failed")
    root.destroy()