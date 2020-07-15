
'''
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
# Alert: internet connection is a must!
# Author: Z.Liang, 20190710PM
# v1: @Z.Liang, added GUI, 20190725PM
# v2: @Z.Liang, 20190726PM, added function which can stop the app running in background
# v3: @Z.Liang, chenged favicon_path
# v4: @Z.liang, fixed an exception for .pyw background retreat; changed property of tmp file -> destroy it after close window
# v5: fix a potential server drive path problem
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
'''

import os, tempfile, distutils.dir_util, shutil, time
from aip import AipOcr
import tkinter
from tkinter import (
    Toplevel,
    PhotoImage,
    Frame,
    Label,
    Entry,
    INSERT,
    Button,
    Text,
    END,
    filedialog,
    messagebox,
)

class App_win:
    # #<~ Baidu's api, personal usage only
    APP_ID = '16753434'
    API_KEY = 'HZD5Dn7Fm6rxFs5z0SXjkkYy'
    SECRET_KEY = 'LQ8GreWFntAGDs5OGuKyYq1FBouWlW7p'
    cmn_width = 40
    excludes = "2.0mm, 0 mm, 02 mm, 0.2mm, mm, 0 mm, mm, m"
    app_width, app_height = 420, 210
    BASE_DIR = os.path.dirname(os.path.abspath(os.path.dirname(__file__)))
    favicon_path = os.path.join(BASE_DIR, r"data\assets\favicon_apple.png")
    tofile = os.path.join(BASE_DIR, r'data\output\dino_data.csv')

    def __init__(self):
        self.tmp = os.path.join(tempfile.gettempdir(), ".{}".format(hash(os.times())))
        if not os.path.isdir(self.tmp): os.makedirs(self.tmp)
        # #<~ top window properties
        self.root = Toplevel()
        self.root.title("Dino OCR v.1, by Z.Liang, 2019")
        try:
            self.imgicon = PhotoImage(file=self.favicon_path)
            self.root.tk.call('wm', 'iconphoto', self.root._w, self.imgicon)
        except tkinter._tkinter.TclError:
            pass
        self.root.geometry("{}x{}".format(self.app_width, self.app_height))
        self.root.focus_set()

        # #<~ frame1
        self.fm1 = Frame(self.root)
        #<~ APP_ID
        self.lbl_id = Label(self.fm1, text="APP_ID: ")
        self.lbl_id.grid(row=0, column=0, sticky='w')

        self.e_id = Entry(self.fm1, width=self.cmn_width)
        self.e_id['show'] = '*'
        self.e_id.grid(row=0, column=1, sticky='w')
        self.e_id.insert(0, self.APP_ID)

        #<~ API_KEY
        self.lbl_key = Label(self.fm1, text="API_KEY: ")
        self.lbl_key.grid(row=1, column=0, sticky='w')

        self.e_key = Entry(self.fm1, width=self.cmn_width)
        self.e_key['show'] = '*'
        self.e_key.grid(row=1, column=1, sticky='w')
        self.e_key.insert(0, self.API_KEY)

        #<~ SECRET_KEY
        self.lbl_secret = Label(self.fm1, text="SECRET_KEY :")
        self.lbl_secret.grid(row=2, column=0, sticky='w')

        self.e_secrect = Entry(self.fm1, width=self.cmn_width)
        self.e_secrect['show'] = '*'
        self.e_secrect.grid(row=2, column=1, sticky='w')
        self.e_secrect.insert(0, self.SECRET_KEY)

        #<~ excludes
        self.lbl_excludes = Label(self.fm1, text="excludes :")
        self.lbl_excludes.grid(row=3, column=0, sticky='nw')

        self.txt_excludes = Text(self.fm1, width=34, height=7)
        self.txt_excludes.config(font=('Arial', 10))
        self.txt_excludes.grid(row=3, column=1, sticky='w')
        self.txt_excludes.insert(INSERT, self.excludes)

        #<~ alert
        self.lbl_alert = Label(self.fm1, text="<<Alert>>:Internet connnection is a must")
        self.lbl_alert.config(font=("Arial", 10))
        self.lbl_alert.grid(row=4, column=1, sticky='w')
        self.fm1.grid(row=0, column=0, sticky='nw')

        # #<~ frame2
        self.fm2 = Frame(self.root)
        #<~ button, click to start ocr
        self.b = Button(self.fm2, text='START', width=8, height=2, command=self.read_dino_data_and_save_as_csv)
        self.b.grid(row=0, column=0, sticky='nw')
        #<~ label, show stats
        self.prg_stat = Label(self.fm2, text='Idle', width=8, height=2)
        self.prg_stat.config(font=('Arial', 10))
        self.prg_stat.grid(row=1, column=0, sticky='nw')
        self.fm2.grid(row=0, column=1, padx=10, sticky='nw')

        # #<~ including a protocol to handle closing on the X button
        self.root.protocol('WM_DELETE_WINDOW', self.close_app)


    def read_dino_data_and_save_as_csv(self):
        '''To read all strings with a specific pattern and store in a list, then return the list'''
        #<~ show script running state: Analyzing
        self.get_prg_stat(stat='Analyzing..', color='orange')
        fd_path = filedialog.askdirectory(title="Choose a folder")
        fd_path = os.path.abspath(fd_path)
        Data = []
        self.txt_excludes.update()
        ex_items = self.get_excludes()
        client = AipOcr(self.get_app_id(), self.get_api_key(), self.get_secrect_key())

        # read all strings from the img folder
        if fd_path != '':
            for dirpath, _, files in os.walk(fd_path):
                for f in files:
                    if f.endswith('.jpg'):
                        f_path = os.path.join(dirpath, f)
                        try:
                            f = open(f_path, 'rb')
                        except IOError:
                            messagebox.showerror("Error","Error: no such file exits")
                        img = f.read()
                        try:
                            msg = client.basicGeneral(img)
                        except:
                            messagebox.showerror("Error", "Error: no internet connection")
                        f = f_path.split(os.path.sep)[-1]
                        txts = [f]
                        try:
                            for w in msg.get('words_result'):
                                txt = w.get('words')
                                if 'm' in txt and not txt in ex_items and not txt.startswith('A'):
                                    if '=' in txt:
                                        txts.append(txt.split('=')[1])
                                    else:
                                        txts.append(txt)
                        except TypeError:
                            messagebox.showerror("Error", "Error: APP_ID / API_ID / SECRET_KEY is NOT available")
                        Data.append(txts)
            #<~ show script running state: Done
            self.get_prg_stat(stat='Done!', color='green')
            # write data into a csv file
            if len(Data):
                tofile_path = os.path.join(self.tmp, self.tofile)
                with open(tofile_path, 'w') as f:
                    for items in Data:
                        for item in items:
                            f.writelines(item + ',')
                        f.writelines('\n')
                # choose a folder to save csv file
                is_yes = messagebox.askquestion("Completed!", "Do you wanna save csv file?")
                if is_yes == 'yes':
                    tofd_path = filedialog.askdirectory(title="Choose a folder")
                    tofd_path = os.path.abspath(tofd_path)
                    if tofd_path != '':
                        distutils.dir_util.copy_tree(self.tmp, tofd_path)
                        messagebox.showinfo("Information", "Your csv file is saved!")
        else:
            messagebox.showerror("Error", "Please choose a folder")
        #<~ show script running state: idle
        self.get_prg_stat()

    def get_app_id(self):
        return self.e_id.get()

    def get_api_key(self):
        return self.e_key.get()

    def get_secrect_key(self):
        return self.e_secrect.get()

    def get_excludes(self):
        if self.excludes != '':
            return [item.strip() for item in self.excludes.split(',')]

    def get_prg_stat(self, stat='Idle', color='black'):
        #<~ show script running state: analyze
        self.prg_stat['text'] = stat
        self.prg_stat['fg'] = color
        time.sleep(.300)

    def mainloop(self):
        self.root.mainloop()

    def close_app(self):
        # #<~ stop app running in background
        shutil.rmtree(self.tmp, ignore_errors=True)
        self.root.destroy()
        self.root.master.deiconify()

def main():
    app = App_win()
    app.mainloop()

if __name__ == "__main__":
    main()
