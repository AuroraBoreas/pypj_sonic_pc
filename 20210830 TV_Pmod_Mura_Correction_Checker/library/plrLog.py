"""
control PLRlog.exe

@ZL, 20210603
"""
import time
import pyautogui
import os
import glob

import subprocess
import contextlib
from typing import NewType
Process = NewType('Process', subprocess.Popen)

@contextlib.contextmanager
def temporary_plrlog(p:Process)->None:
    try:
        yield
    finally:
        p.kill()

class SetLogConverter:
    PLRLOG_OPEN_BTN = (43, 63) # it varies; it depends on PC hardware; and must connect with power not battery;
    EXT             = '.bmp'
    PATTERN         = '*.bin'
    SLEEP_TIME01    = 3
    SLEEP_TIME_INIT = 6
    SLEEP_TIME02    = 2
    SLEEP_TIME03    = 2

    def __init__(self, plr_path: str, bin_folder: str, img_path: str, new_folder: str):
        self.plr_path   = plr_path
        self.bin_files  = [os.path.abspath(p) for p in glob.glob(os.path.join(bin_folder, self.PATTERN))]
        self.img_path   = img_path
        self.new_folder = new_folder

    def to_image(self):
        p = subprocess.Popen(self.plr_path)
        i = 1 # a switch / flag, which makes sure that the first time "open" operation wait a long time; otherwise, wait shorter time;
        with temporary_plrlog(p):
            time.sleep(self.SLEEP_TIME01)
            for bin_file in self.bin_files:
                if i > 1:
                    self.load_MesData_bin(bin_file)
                else:
                    self.load_MesData_bin(bin_file, True)
                i += 1

    def rename_img(self, new_name: str)->None:
        os.rename(self.img_path, os.path.abspath(os.path.join(self.new_folder, new_name + self.EXT)))

    def load_MesData_bin(self, bin_file, isInit: bool = False):
        fw = pyautogui.getActiveWindow()
        fw.maximize() # otherwise PLR GUI position changes per opening
        pyautogui.moveTo(*self.PLRLOG_OPEN_BTN) # move mouse to "open"
        pyautogui.click() # click
        if isInit:
            pyautogui.sleep(self.SLEEP_TIME_INIT)
        else:
            pyautogui.sleep(self.SLEEP_TIME02)
        pyautogui.typewrite(bin_file) # input the absolute path of bin file to address
        pyautogui.sleep(self.SLEEP_TIME03)
        pyautogui.hotkey('Enter') # enter
        pyautogui.sleep(self.SLEEP_TIME03)
        try:
            self.rename_img(os.path.split(bin_file)[-1]) # save img
        except FileExistsError:
            pass

if __name__ == "__main__":
    plr        = r"D:\pj_00_codelib\2019_pypj\20210603 AG85_SET_LED_MURA_Soma\PLRLog.exe"
    src        = r"D:\pj_00_codelib\2019_pypj\20210603 AG85_SET_LED_MURA_Soma\data"
    img_path   = r"D:\pj_00_codelib\2019_pypj\20210603 AG85_SET_LED_MURA_Soma\img.bmp"
    new_folder = r"D:\pj_00_codelib\2019_pypj\20210603 AG85_SET_LED_MURA_Soma\Images"
    
    slc = SetLogConverter(plr, src, img_path, new_folder)
    slc.to_image()