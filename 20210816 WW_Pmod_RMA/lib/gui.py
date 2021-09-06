"""
A simple UI class gets user input.

@ZL, 20210804

"""

import tkinter
import sys
import os, datetime
from tkinter import filedialog

sys.path.append('.')
from lib.quotation import Tracker, logging
from lib.bom import generate_bom_excel
from lib.mail import Mailer
from lib import config

class BomTracker:
    _w, _h           = 400, 250
    _entry_width     = 30
    _btn_start_width = 20
    _xpadding        = 25
    _btn_width       = 8

    def __init__(self, tracker:Tracker):
        self._tracker = tracker
        self.root     = tkinter.Toplevel()
        self.root.title("BOM Buyback quotation Tracker v0.01 @ZL")

        self._ws, self._hs = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        x = int((self._ws - self._w) / 2)
        y = int((self._hs - self._h) / 2)
        self.root.geometry("{0}x{1}+{2}+{3}".format(self._w, self._h, x, y))
        self.root.focus_set()

        self.fm = tkinter.Frame(self.root)
        self.fm.grid(row=0, column=0, padx=self._xpadding, pady=self._btn_width, sticky=tkinter.NW)

        self.lbl_Create = tkinter.Label(self.fm, text="Create(FromExcel) :")
        self.lbl_Create.grid(row=0, column=0, sticky=tkinter.W)
        
        self.lbl_Export = tkinter.Label(self.fm, text="Export(ToExcel) :")
        self.lbl_Export.grid(row=1, column=0, sticky=tkinter.W)

        self.lbl_Import = tkinter.Label(self.fm, text="Import(FromFolder) :")
        self.lbl_Import.grid(row=2, column=0, sticky=tkinter.W)

        self.lbl_Generate = tkinter.Label(self.fm, text="Generate(BOM Excel) :")
        self.lbl_Generate.grid(row=3, column=0, sticky=tkinter.W)

        self.lbl_Update = tkinter.Label(self.fm, text="Update(BOM status) :")
        self.lbl_Update.grid(row=4, column=0, sticky=tkinter.W)

        self.btn_Create = tkinter.Button(self.fm, width=self._btn_start_width, text="Create", command=self.create)
        self.btn_Create.grid(row=0, column=1, pady=self._btn_width, sticky=tkinter.NW)

        self.btn_Export = tkinter.Button(self.fm, width=self._btn_start_width, text="Export", command=self.export)
        self.btn_Export.grid(row=1, column=1, pady=self._btn_width, sticky=tkinter.NW)

        self.btn_Import = tkinter.Button(self.fm, width=self._btn_start_width, text="Import", command=self.emport)
        self.btn_Import.grid(row=2, column=1, pady=self._btn_width, sticky=tkinter.NW)

        self.btn_Generate = tkinter.Button(self.fm, width=self._btn_start_width, text="Generate", command=self.generate)
        self.btn_Generate.grid(row=3, column=1, pady=self._btn_width, sticky=tkinter.NW)

        self.btn_Update = tkinter.Button(self.fm, width=self._btn_start_width, text="Update", command=self.update)
        self.btn_Update.grid(row=4, column=1, pady=self._btn_width, sticky=tkinter.NW)

        # self.fm2 = tkinter.Frame(self.root)
        # self.fm.grid(row=1, column=0, padx=self._xpadding, pady=self._btn_width, sticky=tkinter.NW)
        # self.txt_Logger = tkinter.Text(self.fm2, width=140, height=48)
        # self.txt_Logger.grid(row=0, column=0, sticky="W")
        # self.txt_Logger.insert(tkinter.INSERT, "")

        self.root.protocol("WM_DELETE_WINDOW", self.close)

    def create(self)->None:
        # @TODO: create table based on an existing excel workbook
        src_path = filedialog.askopenfilename(filetypes=(('Excel files', '*.xls*'), ('all files', '*.*')))
        if os.path.exists(src_path):
            self._tracker.create(src_path)

    def export(self)->None:
        # @TODO: export to excel
        dst_folder = filedialog.askdirectory()
        dst = os.path.join(dst_folder, 'bom_quotation.xlsx')
        if os.path.exists(dst_folder):
            self._tracker.export(dst)

    def emport(self)->None:
        # @TODO: append data from new lcm_ww_quotation files into sqlite database recursively
        src_folder = filedialog.askdirectory()
        ibm = 0
        bmd = 'TBD'
        if os.path.exists(src_folder):
            for root, _, files in os.walk(src_folder):
                for file in files:
                    self._tracker.emport(os.path.join(root, file), is_bom_maintained=ibm, bom_maintain_date=bmd)
    
    def generate(self)->None:
        dst_folder = filedialog.askdirectory()
        today      = datetime.datetime.today().strftime("%Y%m%d")
        fname      = '联络书（新品维护）ver2 210730_{0}.xlsx'.format(today)
        dst        = os.path.abspath(os.path.join(dst_folder, fname))

        if os.path.exists(dst_folder):
            # @TODO: query wait-to-register quotations
            rv = self._tracker.retrieve(is_bom_maintained=0)

            # @TODO: write them into bom-registeration-excel
            df = rv[['PMOD', 'BuybackProductNo', 'Price', 'ModelName']]
            df_bom = df.assign(
                状态=r'部品',
                维护到=r'ASSY'
            )
        
            if len(df_bom) > 0:
                try:
                    generate_bom_excel(config.tmp_bom, df_bom, dst)
                    
                    subject = "新品维护 {0}".format(today)
                    body = """
                    <p><font face="DengXian">Jin san，
                    <br>你好
                    <br>
                    <br>附件帮忙做一下新品维护。
                    <br>谢谢
                    <br></font></p>    
                    """
                    mailer = Mailer(subject, body, dst)
                    mailer.send()
                except AttributeError:
                    logging.info('Failed: generate_bom_excel')
            else:
                logging.info('0 BuybackProductNo wait-to-register!')

    def update(self)->None:
        now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self._tracker.update(0, 1, now)

    def close(self)->None:
        self.root.destroy()
        self.root.master.deiconify()
    
    def loop(self)->None:
        self.root.mainloop()

def main():
    tracker = Tracker(config.db)
    app     = BomTracker(tracker)
    app.loop()

if __name__ == '__main__':
    main()