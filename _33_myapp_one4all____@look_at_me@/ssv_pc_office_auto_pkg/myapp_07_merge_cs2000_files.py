#!/usr/bin/env python
# coding: utf-8

"""
================================================================================
A tool merges all cs2000 optical files based on file created date.
@ZL, 20191106
v0, fix a potential server drive path problem
v1, @ZL, 20200603, fixed a bug
================================================================================
"""

import os, glob, collections, distutils.dir_util, tempfile
import pandas as pd
import tkinter as tk

from tkinter import filedialog
from datetime import datetime

def main():
    root = tk.Toplevel()
    root.withdraw()
    # tmp = os.path.join(tempfile.gettempdir(), '.{}'.format(hash(os.times())))
    # if not os.path.isdir(tmp): os.makedirs(tmp)

    fd_path = filedialog.askdirectory()
    fd_path = os.path.abspath(fd_path)
    # fd_path = r"C:\Users\5106001995\Desktop\cs2000_files merge demo\20180817 SB49&55&65 CS Drift\20180817 SB55&65 CS Drift"
    if os.path.isdir(fd_path):
        files = glob.glob(os.path.join(fd_path, "*.xlsx"))
        if files:
            #<~ sort file names based on its created date
            d = {}
            for f in files:
                t = os.path.getmtime(f)
                d.setdefault(t, f)
            od = collections.OrderedDict(sorted(d.items(), key=lambda t: t[0]))

            #<~ open file, read data, add filename col
            df_li = []
            for t, f in od.items():
                df = pd.read_excel(f)
                df['Lv'] = df['L[cd/m^2]']
                df['file_name'] = os.path.split(f)[-1]
                df['modified_time'] = datetime.fromtimestamp(t).strftime('%Y-%m-%d %H:%M:%S')
                df.drop(columns=['L[cd/m^2]'], inplace=True)
                df_li.append(df)

            #<~ combine all date and save as csv file
            big_df = pd.concat(df_li, ignore_index=True)
            big_df.to_csv(os.path.join(fd_path, 'all_merged.csv'))

        root.destroy()
        root.master.deiconify()

if __name__ == '__main__':
    main()