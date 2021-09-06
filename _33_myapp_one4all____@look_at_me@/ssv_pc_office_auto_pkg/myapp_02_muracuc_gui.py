'''
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
# This is an alternative way to extract data from .eep file and plot image
# author: Z.Liang, 20190505
# v2, @Z.Liang, 20190725. GUI added, temp folder and save images
# v3, @Z.Liang, 20190726, restructed
# v4: @Z.liang, fixed an exception for .pyw background retreat
# v5: @Z.liang, 20190806, updates as follows.
      utilize pandas to read src file instead of normal method, open(f, read_mode='r');
      get rid of shutil.rm_tree(dir, ignore_errors=True) after read an article on https://www.codercto.com/a/41423.html;
# v6: fix a potential server drive path problem
# v7: bugfix, @ZL, 20201207
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
'''

import os, tempfile, distutils.dir_util, time
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import tkinter
from tkinter import (
    Toplevel,
    PhotoImage,
    Label,
    INSERT,
    Button,
    Text,
    END,
    filedialog,
    messagebox,
)
from PIL import (
    ImageTk, Image
)

class App_win():
    BASE_DIR = os.path.dirname(os.path.abspath(os.path.dirname(__file__)))
    cmn_width = 26
    ffamily = "Arial Narrow"
    fsize = 10
    favicon_path = os.path.join(BASE_DIR, r"data\assets\favicon_mushroom.png")

    def __init__(self):
        # #<~ toplevel
        self.tmp = os.path.join(tempfile.gettempdir(), '.{}'.format(hash(os.times())))
        if not os.path.isdir(self.tmp): os.makedirs(self.tmp)

        self.root = Toplevel()
        self.root.title("CUC Sim v.1, by Z.Liang, 2019")
        self.root.geometry("{}x{}".format(360, 620))
        try:
            self.imgicon = PhotoImage(file=self.favicon_path)
            self.root.tk.call('wm', 'iconphoto', self.root._w, self.imgicon)
        except tkinter._tkinter.TclError:
            pass
        self.root.focus_set()

        # #<~ label to show app stat
        self.app_stat = Label(self.root, text='idle', width=8, height=1)
        self.app_stat.config(font=(self.ffamily, 10, 'bold'))
        self.app_stat.grid(row=0, column=1, sticky="N")

        # #<~ label to display image
        self.panel_img = Label(self.root, text="image will display here", image=None)
        self.panel_img.config(font=(self.ffamily, self.fsize))
        self.panel_img.grid(row=3, column=0, sticky="w")

        # #<~ text to display panel SN
        self.panel_sn = Text(self.root, width=57, height=1)
        self.panel_sn.config(font=(self.ffamily, 9))
        self.panel_sn.insert(INSERT, "panel S/N:")
        self.panel_sn.grid(row=2, column=0, sticky="w")

        # #<~ button to choose folder where images could be stored
        self.img_fd = Button(self.root, text="Choose a folder to save images", width=self.cmn_width, anchor="w", command=self.save_imgs_into_fd)  #<~logic
        self.img_fd.config(font=(self.ffamily, 12))
        self.img_fd.grid(row=1, column=0, sticky="w")

        # #<~ button to analyze eep file
        self.b = Button(self.root, text='Choose an eep file', width=self.cmn_width, anchor="w", command=self.extract_zvalue_and_plot)    #<~logic
        self.b.config(font=(self.ffamily, 12))
        self.b.grid(row=0, column=0, sticky="w")

        # #<~ stop the process running in background
        self.root.protocol("WM_DELETE_WINDOW", self.close_app)

    def find_index(self, r, c, color='RED'):
        """
        Purpose:Return calculation result(index) based on color
        author:Z.Liang, 20190505
        """
        if color == 'RED':
            return int('6019', 16) - int('6000', 16) + c * 36 + r + 1 - (r % 2) * 2
        elif color == 'GREEN':
            return int('6919', 16) - int('6000', 16) + c * 36 + r + 1 - (r % 2) * 2
        elif color == 'BLUE':
            return int('7219', 16) - int('6000', 16) + c * 36 + r + 1 - (r % 2) * 2

    def make_colormap(self, seq):
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

    def extract_zvalue_and_plot(self):
        """
        Purpose: save files to a fold with a given name on any user's desktop
        author: Z.Liang, 20190505
        v1, 20190719: added GUI design
        v2, 20190725: added error-proofing for filedialog.askopenfilename() when it doesnt return a valid path
        v3, 20201207: bugfix
        """
        # #<~ show app stat: analyzing
        self.change_app_stat(text='Analyzing..', color='orange')
        f_path = filedialog.askopenfilename(filetypes=(('eep files', '*.eep*'), ('all files', '*.*')))

        if f_path != '':
            save2array = np.array([0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0])
            ## extract src data
            df = pd.read_csv(f_path, header=None, skiprows=7, sep=r'\t|\s+', engine='python').drop([0], axis=1)
            save2array = np.vstack((save2array, df.values.tolist()))
            b = save2array.flatten()[:6828]

            M, N = 36, 61
            colors = ['RED', 'GREEN', 'BLUE']
            c = mcolors.ColorConverter().to_rgb
            ryg = self.make_colormap([c('#F7696B'), c('#ffeb84'), 0.50, c('#ffeb84'), c('#CEDC81'), 0.77, c('#CEDC81'), c('#63be7b')])
            fn = os.path.split(f_path)[1]

            ## plot data for each color
            subplot_datas = []
            for color in colors:
                arrZvalue = (np.array([[int(b[self.find_index(r, c, color=color)-1], 16)
                            if int(b[self.find_index(r, c, color=color)-1], 16) < 127
                            else int(b[self.find_index(r, c, color=color)-1], 16)-256]
                            for r in range(0, M) for c in range(0, N)]).flatten().reshape(M, N))
                subplot_datas.append(arrZvalue)

            ## plot
            i = 1
            subplot_titles = colors
            subplot_colors = [self.make_colormap([c('#F7696B')]), self.make_colormap([c('#63be7b')]), self.make_colormap([c('blue')])]
            for subplot_title, subplot_color, subplot_data in zip(subplot_titles, subplot_colors, subplot_datas):
                if not np.all(subplot_data==0):
                    _cmap = ryg
                else:
                    _cmap = subplot_color
                plt.subplot(3, 1, i)
                plt.imshow(subplot_data, cmap=_cmap, interpolation='nearest')
                plt.axis("off")
                plt.title(f'{subplot_title}', fontsize=12)
                i += 1

            ##<~ change font size in the plt
            medium_size = 10
            plt.rc('xtick', labelsize=medium_size)
            plt.rc('ytick', labelsize=medium_size)
            plt.subplots_adjust(left=None, bottom=0.1, right=None, top=1.8, wspace=0.2, hspace=0.2)

            ##<~ save to the temp folder
            tmp_fn = os.path.join(self.tmp, f'{fn}.png')
            plt.savefig(tmp_fn, dpi=300, bbox_inches='tight')
            plt.clf()

            # #<~ display image on app window
            width, height = 287, 518
            img = Image.open(tmp_fn)
            img = img.resize((width, height), Image.ANTIALIAS)
            photoimg = ImageTk.PhotoImage(img)
            self.panel_img.configure(image=photoimg)
            self.panel_img.image = photoimg
            self.panel_sn.delete(1.0, END)
            self.panel_sn.insert(INSERT, fn)
            # #<~ show app stat: Done
            self.change_app_stat(text='Done!', color='green')
        else:
            # #<~ show app stat: idle
            self.change_app_stat()

    def save_imgs_into_fd(self):
        fd_path = filedialog.askdirectory(title='Choose Directory to save')
        fd_path = os.path.abspath(fd_path)
        if fd_path != '':
            #<~ save imgs into a temp folder
            distutils.dir_util.copy_tree(self.tmp, fd_path)
            messagebox.showinfo("Information", "Saved OK!")
            # webbrowser.open('file:///' + fd_path)

    def change_app_stat(self, text='idle', color='black'):
        self.app_stat['text'] = text
        self.app_stat['fg'] = color
        time.sleep(.5)

    def close_app(self):
        distutils.dir_util.remove_tree(self.tmp)
        self.root.destroy()
        self.root.master.deiconify()

    def mainloop(self):
        self.root.mainloop()

def main():
    app = App_win()
    app.mainloop()

if __name__ == '__main__':
    main()
