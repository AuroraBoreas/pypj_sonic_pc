"""
A simple UI class gets user input.

v0.01, @ZL, 20210506
v0.02, @ZL, 20210820
"""

import tkinter
import sys

sys.path.append('.')
from ssv_pc_office_auto_pkg.nmi.template import PowerPointTemplate

class ReportTemplate:
    _w, _h = 500, 240
    _entry_width = 30
    _btn_start_width = 20
    _xpadding = 25

    def __init__(self, *args):
        """
        args[0] --> ssve_template;
        args[1] --> save_path;
        """
        self.args = args
        self.root = tkinter.Toplevel()
        self.root.title("NMI Reporter v.0.1 @ZL")

        self._ws, self._hs = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        x = int((self._ws - self._w) / 2)
        y = int((self._hs - self._h) / 2)
        self.root.geometry("{0}x{1}+{2}+{3}".format(self._w, self._h, x, y))
        self.root.focus_set()

        self.fm = tkinter.Frame(self.root)
        self.fm.grid(row=0, column=0, padx=self._xpadding, pady=10, sticky=tkinter.NW)

        self.lbl_Pmod = tkinter.Label(self.fm, text="PMod/ITC :")
        self.lbl_Pmod.grid(row=0, column=0, sticky=tkinter.W)

        self.lbl_Symptom = tkinter.Label(self.fm, text="Symptom :")
        self.lbl_Symptom.grid(row=1, column=0, sticky=tkinter.W)

        self.lbl_Locality = tkinter.Label(self.fm, text="Locality :")
        self.lbl_Locality.grid(row=2, column=0, sticky=tkinter.W)

        self.lbl_Event = tkinter.Label(self.fm, text="Event :")
        self.lbl_Event.grid(row=3, column=0, sticky=tkinter.W)

        self.lbl_Copies = tkinter.Label(self.fm, text="Copies :")
        self.lbl_Copies.grid(row=4, column=0, sticky=tkinter.W)

        self.e_Pmod = tkinter.Entry(self.fm, width=self._entry_width)
        self.e_Pmod.insert(tkinter.INSERT, "AG")
        self.e_Pmod.grid(row=0, column=1, sticky=tkinter.W)

        self.e_Symptom = tkinter.Entry(self.fm, width=self._entry_width)
        self.e_Symptom.insert(tkinter.INSERT, "V-line")
        self.e_Symptom.grid(row=1, column=1, sticky=tkinter.W)

        self.e_Locality = tkinter.Entry(self.fm, width=self._entry_width)
        self.e_Locality.insert(tkinter.INSERT, "SOEM")
        self.e_Locality.grid(row=2, column=1, sticky=tkinter.W)

        self.e_Event = tkinter.Entry(self.fm, width=self._entry_width)
        self.e_Event.insert(tkinter.INSERT, "MP")
        self.e_Event.grid(row=3, column=1, sticky=tkinter.W)

        self.e_Copies = tkinter.Entry(self.fm, width=self._entry_width)
        self.e_Copies.insert(tkinter.INSERT, "1")
        self.e_Copies.grid(row=4, column=1, sticky=tkinter.W)

        self.btn_Start = tkinter.Button(self.fm, width=self._btn_start_width, text="Start", command=self.create)
        self.btn_Start.grid(row=5, column=1, pady=10, sticky=tkinter.NW)

        self.root.protocol("WM_DELETE_WINDOW", self.close)

    def create(self):
        ppt_template = PowerPointTemplate(
            *self.args,
            self.e_Pmod.get(),
            self.e_Symptom.get(),
            self.e_Locality.get(),
            self.e_Event.get(),
            self.e_Copies.get()
        )
        ppt_template.build()

    def close(self):
        self.root.destroy()
        self.root.master.deiconify()
    
    def loop(self):
        self.root.mainloop()

def main():
    tmp_path  = r'D:\pj_00_codelib\2019_pypj\_33_myapp_one4all____@look_at_me@\data\templates\template.pptx'
    save_path = r"C:\Users\5106001995\Desktop"
    rpt_tmp   = ReportTemplate(tmp_path, save_path)
    rpt_tmp.loop()

if __name__ == '__main__':
    main()