"""
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
all in one
v0.00: @Z.Liang, 20190802
v0.01: @ZL, 20190819. added random color for background of GUI
v0.02: @Z.Liang, 20190905, added optical report
v0.03: @ZL, 20191022, orgnized modules of myapps
v0.04: @ZL, 20200522, standalone packages
v0.05: @ZL, 20210506, add NMI reporter
v0.06: @ZL, 20210720, add ssv query
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
from ssv_pc_office_auto_pkg.nmi import gui as nmi_gui
from ssv_pc_office_auto_pkg.ulpayment import gui as ulpayment_gui
from ssv_pc_office_auto_pkg.ssve_server import ssve_qry

import tkinter, os
from tkinter import (
    Frame,
    PhotoImage,
    Button,
)

class app_wm:
    btn_width = 30
    BASE_DIR = os.path.dirname(__file__)
    favicon_path = os.path.join(BASE_DIR, r"data\assets\favicon_sword.png")
    color_map = ['azure', 'cyan', 'spring green', 'yellow', 'salmon']

    def __init__(self):
        self.root = tkinter.Tk()
        self.root.title("My Apps v2 @ZL")
        self.root.config(bg=self.color_map[0])

        #get screen width and height
        self.w, self.h = 680, 360
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
        #<~myapp09: nmi ppt template
        self.btn_nmi_reporter = Button(self.fm2, text='10. NMI reporter', width=self.btn_width, anchor='w', command=self.display_nmi_reporter)
        self.btn_nmi_reporter.grid(row=3, column=1, pady=2, sticky='w')
        #<~myapp10: ul payment
        self.btn_ul_payment = Button(self.fm2, text='11. UL py Mail', width=self.btn_width, anchor='w', command=self.display_ul_payment)
        self.btn_ul_payment.grid(row=4, column=1, pady=2, sticky='w')
        #<~myapp11: ssv query
        self.btn_ssv_query = Button(self.fm2, text='12. ssv query', width=self.btn_width, anchor='w', command=self.display_ssv_query)
        self.btn_ssv_query.grid(row=5, column=1, pady=2, sticky='w')

        self.root.protocol("WM_DELETE_WINDOW", self.close_app)

    def close_app(self):
        self.root.destroy()

    def mainloop(self):
        self.root.mainloop()

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

    def display_nmi_reporter(self):
        self.root.withdraw()
        nmi_gui.main()

    def display_ul_payment(self):
        self.root.withdraw()
        ulpayment_gui.main()

    def display_ssv_query(self):
        os.startfile(ssve_qry)

def main():
    app = app_wm()
    app.mainloop()

if __name__ == "__main__":
    main()