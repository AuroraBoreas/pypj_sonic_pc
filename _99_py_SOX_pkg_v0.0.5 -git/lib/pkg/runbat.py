"""A simple module runs batch commands, a basic wrapper

This module has basic functionalities as follows.
- passes gain_value to adb CLI, change TV src_rms value
- runs batches file and waits till done

Changelog
- v0.0.1, initial version
- v0.0.2, bugfix

Author
- @ZL, 20200629
- @SHES
"""

import subprocess

def change_src_gain(gain_value):
    """passes *gain_value* to adb CLI, change TV src_rms value
    Author: SHES

    :param gain_value: gain value
    :type gain_value: float
    """
    subprocess.call("adb root", shell=True)
    subprocess.call("adb shell setenforce 0", shell=True)
    subprocess.call("adb shell chmod 777 data", shell=True)
    # change src gain
    subprocess.call("adb shell setprop vendor.mtk.audio.aec.ref.gain {}".format(gain_value), shell=True)

def run_batch(batch_file):
    """wrapper to run batch file

    :param batch_file: name or path of batch file
    :type batch_file: string
    """
    proc = subprocess.Popen(batch_file, shell=True)
    proc.wait()