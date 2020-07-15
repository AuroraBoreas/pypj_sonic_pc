'''
====================================================================================
This is an alternative way to extract data from .eep file and plot image
author: Z.Liang, 20190430
v1.mutiprocesses to improve performance, @ZL, 20190925
v1.1logger to record history, @ZL, 20190925
v2. fix a potential server drive path problem
====================================================================================
'''

import os, getpass, itertools, glob, time, concurrent.futures
import tkinter as tk
from tkinter import filedialog
# from tkinter import TopLevel
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors

def find_index(r, c, color='RED'):
    """
    Purpose:Return calculation result(index) based on color
    author: Z.Liang, 20190505
    """
    if color == 'RED':
        return int('6019', 16) - int('6000', 16) + c * 36 + r + 1 - (r % 2) * 2
    elif color == 'GREEN':
        return int('6919', 16) - int('6000', 16) + c * 36 + r + 1 - (r % 2) * 2
    elif color == 'BLUE':
        return int('7219', 16) - int('6000', 16) + c * 36 + r + 1 - (r % 2) * 2

def make_colormap(seq):
    """
    Return a LinearSegmentedColormap
    seq: a sequence of floats and RGB-tuples. The floats should be increasing and in the interval (0,1)
    Author: SOF
    """
    seq = [(None,) * 3, 0.0] + list(seq) + [1.0, (None,) * 3]
    cdict = {'red': [], 'green': [], 'blue': []}
    for i, item in enumerate(seq):
        if isinstance(item, float):
            r1, g1, b1 = seq[i - 1]
            r2, g2, b2 = seq[i + 1]
            cdict['red'].append([item, r1, r2])
            cdict['green'].append([item, g1, g2])
            cdict['blue'].append([item, b1, b2])
    return mcolors.LinearSegmentedColormap('CustomMap', cdict)

def extract_zvalue_and_plot(f_path, save2fdname='MURACUC_Img'):
    """
    Purpose: save files to a fold with a given name on any user's desktop
    Author: Z.Liang, 20190505
    mutiprocesses to improve performance, @ZL, 20190925
    """
    save2fd_path = os.path.join("C:\\Users", getpass.getuser(), "Desktop", save2fdname)
    if not os.path.isdir(save2fd_path):
        os.makedirs(save2fd_path, exist_ok=True)

    # extract src data
    df = pd.read_csv(f_path, header=None, skiprows=7, sep=r'\t|\s+', engine='python')
    dummy = np.array([0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0])
    new_df = df.drop([0], axis=1)
    new_df = np.vstack((dummy, new_df.values.tolist()))
    b = new_df.flatten()[:6828]

    M, N = 36, 61
    colors = ['RED', 'GREEN', 'BLUE']
    c = mcolors.ColorConverter().to_rgb
    ryg = make_colormap([c('red'), c('#ffeb84'), 0.38, c('#ffeb84'), c('#63be7b'), 0.77, c('#63be7b')])

    fn = os.path.split(f_path)[1] ##<~to: A2231641A_0000434
    # plot data for each color
    i = 1
    for color in colors:
        #  hexdata: the hardest part
        arrZvalue = (np.array([[int(b[find_index(r, c, color=color)-1], 16) if int(b[find_index(r, c, color=color)-1], 16) < 127 else
                                int(b[find_index(r, c, color=color)-1], 16)-256] for r in range(0, M) for c in range(0, N)]).flatten().reshape(M, N))

        plt.subplot(3, 1, i)
        plt.imshow(arrZvalue, cmap=ryg, interpolation='nearest')
        plt.title(f'{color}', fontsize=10)
        i += 1
    # save to a given folder
    plt.subplots_adjust(left=None, bottom=0.1, right=None, top=1.8, wspace=0.2, hspace=0.2)
    plt.savefig(f'{save2fd_path}/{fn}.png', dpi=300, bbox_inches='tight')
    plt.clf()

def main():
    """Batch codes here"""
    root = tk.Toplevel()
    root.withdraw()
    # root.title = "MURACUC Img Simulator v.1"
    fd_path = filedialog.askdirectory()
    fd_path = os.path.abspath(fd_path)
    if fd_path != '':
        all_files = glob.glob(os.path.join(fd_path, '*.eep'))

        #<~~ mutiprocesses to improve performance, @ZL, 20190925
        with concurrent.futures.ProcessPoolExecutor() as executor:
            executor.map(extract_zvalue_and_plot, all_files)

        root.destroy()
        root.master.deiconify()

if __name__ == '__main__':
    main()
