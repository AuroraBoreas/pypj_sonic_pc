"""A simple module runs batch commands, a basic wrapper

"""

import subprocess

def change_src_gain(gain_value):
    """Author: SHES

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