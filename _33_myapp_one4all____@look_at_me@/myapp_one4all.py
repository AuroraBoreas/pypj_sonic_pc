"""
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
# all in one
# v0: @Z.Liang, 20190802
# v1: @ZL, 20190819. added random color for background of GUI
# V2: @Z.Liang, 20190905, added optical report
# V3: @ZL, 20191022, orgnized modules of myapps
# V4: @ZL, 20200522, standalone packages
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
"""
from ssv_pc_office_auto_pkg import myapp_00_muracuc_batch as muracuc_batch
from ssv_pc_office_auto_pkg import myapp_01_ul_contract as UL
from ssv_pc_office_auto_pkg import myapp_02_muracuc_gui as muracuc_gui
from ssv_pc_office_auto_pkg import myapp_03_Dino_ocr as dino
from ssv_pc_office_auto_pkg import myapp_04_Optical_report as optical_report
from ssv_pc_office_auto_pkg import myapp_05_FOS_report as FOS_report
from ssv_pc_office_auto_pkg import myapp_06_upload_email as upload_email
from ssv_pc_office_auto_pkg import myapp_07_merge_cs2000_files as cs2000_files_merger
from ssv_pc_office_auto_pkg import myapp_08_kill_tmp_folders as tmp_dirs_killer

import tkinter, os
from tkinter import (
    Frame,
    PhotoImage,
    Button,
)

class app_wm():
    btn_width = 25
    BASE_DIR = os.path.dirname(__file__)
    favicon_path = os.path.join(BASE_DIR, r"data\assets\favicon_sword.png")
    color_map = ['azure', 'cyan', 'spring green', 'yellow', 'salmon']

    def __init__(self):
        self.root = tkinter.Tk()
        self.root.title("My Apps v1, @Z.Liang | 2019")
        self.root.config(bg=self.color_map[0])

        #get screen width and height
        self.w, self.h = 420, 200
        self.ws = self.root.winfo_screenwidth()
        self.hs = self.root.winfo_screenheight()
        x = int((self.ws - self.w) / 2)
        y = int((self.hs - self.h) / 2)
        self.root.geometry("{}x{}+{}+{}".format(self.w, self.h, x, y))
        self.imgicon = PhotoImage(file=self.favicon_path)
        self.root.tk.call("wm", "iconphoto", self.root._w, self.imgicon)

        ##<~frame1
        self.fm1 = Frame(self.root)
        self.fm1.grid(row=0, column=0, sticky='w', padx=10, pady=5)
        #<~myapp00: UL contract
        self.btn_ul = Button(self.fm1, text='01. UL contract generator', width=self.btn_width, anchor='w', command=self.display_ul)
        self.btn_ul.grid(row=0, column=0, pady=2, sticky='w')
        #<~myapp01: Dino OCR
        self.btn_dino = Button(self.fm1, text='02. Dino OCR', width=self.btn_width, anchor='w', command=self.display_dino)
        self.btn_dino.grid(row=1, column=0, pady=2, sticky='w')
        #<~myapp02: MURACUC batch
        self.btn_mcbatch = Button(self.fm1, text='03. MURACUC batch converter', width=self.btn_width, anchor='w', command=self.display_muracuc_batch)
        self.btn_mcbatch.grid(row=2, column=0, pady=2, sticky='w')
        #<~myapp03: MURACUC GUI
        self.btn_mcgui = Button(self.fm1, text='04. MURACUC Gui', width=self.btn_width, anchor='w', command=self.display_muracuc_gui)
        self.btn_mcgui.grid(row=3, column=0, pady=2, sticky='w')
        #<~myapp04: Optical report
        self.btn_optical_report = Button(self.fm1, text='05. Optical Report', width=self.btn_width, anchor='w', command=self.display_optical_report_gui)
        self.btn_optical_report.grid(row=4, column=0, pady=2, sticky='w')
        #<~myapp05: FOS report
        self.btn_FOS_report = Button(self.fm1, text='06. FOS Report', width=self.btn_width, anchor='w', command=self.display_FOS_report_gui)
        self.btn_FOS_report.grid(row=5, column=0, pady=2, sticky='w')


        ##<~frame2
        self.fm2 = Frame(self.root)
        self.fm2.grid(row=0, column=1, sticky='n', padx=10, pady=5)
        #<~myapp06: upload to server and email
        self.btn_upload_n_email = Button(self.fm2, text='07. Upload & send email', width=self.btn_width, anchor='w', command=self.display_upload_n_email_gui)
        self.btn_upload_n_email.grid(row=0, column=1, pady=2, sticky='w')
        #<~myapp07: merge cs2000 files
        self.btn_cs2000_files_merger = Button(self.fm2, text='08. merge cs2000 files', width=self.btn_width, anchor='w', command=self.display_cs2000_files_merger)
        self.btn_cs2000_files_merger.grid(row=1, column=1, pady=2, sticky='w')
        #<~myapp08: merge cs2000 files
        self.btn_tmp_dirs_killer = Button(self.fm2, text='09. Clear temp folders in AppData', width=self.btn_width, anchor='w', command=self.display_tmp_dirs_killer)
        self.btn_tmp_dirs_killer.grid(row=2, column=1, pady=2, sticky='w')


        self.root.protocol("WM_DELETE_WINDOW", self.close_app)

    def display_ul(self):
        self.root.withdraw()
        UL.main()

    def display_dino(self):
        self.root.withdraw()
        dino.main()

    def display_muracuc_batch(self):
        muracuc_batch.main()

    def display_muracuc_gui(self):
        self.root.withdraw()
        muracuc_gui.main()

    def display_optical_report_gui(self):
        self.root.withdraw()
        optical_report.main()

    def display_FOS_report_gui(self):
        self.root.withdraw()
        FOS_report.main()

    def display_upload_n_email_gui(self):
        self.root.withdraw()
        upload_email.main()

    def display_cs2000_files_merger(self):
        self.root.withdraw()
        cs2000_files_merger.main()

    def display_tmp_dirs_killer(self):
        # self.root.withdraw() #<~ no GUI in this module
        tmp_dirs_killer.main()

    def close_app(self):
        self.root.destroy()

    def mainloop(self):
        self.root.mainloop()

def main():
    app = app_wm()
    app.mainloop()

if __name__ == "__main__":
    main()
