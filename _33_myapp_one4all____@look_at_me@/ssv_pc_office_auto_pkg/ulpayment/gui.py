"""
============================================================================================================================
#Z.Liang, 20191031
#ZL, 20210518

Changelog
#v0, make life easier: to upload to server and send email automatically.
#v1, fix a potential server drive path problem
#v2, add slide images into email body to avoid manual copy-paste
============================================================================================================================
"""

import os, sys, tempfile, threading, socket, distutils.dir_util, time
import win32com.client as win32
from distutils.dir_util import copy_tree

import tkinter
from tkinter.ttk import Progressbar
from tkinter import (
    Toplevel,
    PhotoImage,
    Frame,
    Label,
    Entry,
    INSERT,
    Button,
    NW,
    END,
    HORIZONTAL,
    filedialog,
    messagebox,
)

sys.path.append('.')
from ssv_pc_office_auto_pkg.ulpayment import mail

class App_win():
    upload2server_fd_path = r"\\43.98.1.18\SSVE_Division\SHES-C\部共通\02-Sections\4020_Panel_Design\经费管理\system application report\决裁&入力"
    server_address = '43.98.1.18'
    port = 445

    ##<~ tmp folder here is NOT for this module. it's for integral GUI<all4one>.
    tmp = os.path.join(tempfile.gettempdir(), '.{}'.format(hash(os.times())))
    if not os.path.isdir(tmp): os.makedirs(tmp)

    def __init__(self):
        """===GUI==="""
        self.root = Toplevel()  ##<~ toplevel when import from other module
        self.root.title("ToServer & EMail v.1 @ZL")

        self.w, self.h = 1260, 240
        self.ws = self.root.winfo_screenwidth()
        self.hs = self.root.winfo_screenheight()
        x = int((self.ws - self.w) / 2)
        y = int((self.hs - self.h) / 2)
        self.root.geometry("{}x{}+{}+{}".format(self.w, self.h, x, y))
        self.root.focus_set()

        ##<~frame1
        self.fm1 = Frame(self.root)
        self.fm1.grid(row=0, column=0, padx=2, sticky='nw')

        self.lbl_ppt_report = Label(self.fm1, text='Payment file:')
        self.lbl_ppt_report.grid(row=0, column=0, sticky='w')
        self.lbl_upload2server = Label(self.fm1, text='Server:')
        self.lbl_upload2server.grid(row=1, column=0, sticky='w')

        self.e_ppt_report = Entry(self.fm1, width=100)
        self.e_ppt_report.grid(row=0, column=1, sticky='w')
        self.e_upload2server = Entry(self.fm1, width=100)
        self.e_upload2server.insert(INSERT, self.upload2server_fd_path)
        self.e_upload2server.grid(row=1, column=1, sticky='w')

        self.btn_add_report = Button(self.fm1, text='+', command=self.pick_ppt_report_file_path_dialog)
        self.btn_add_report.grid(row=0, column=2, padx=1, pady=1, sticky='nw')
        self.btn_upload2server = Button(self.fm1, text='+', command=self.pick_up2server_path_dialog)
        self.btn_upload2server.grid(row=1, column=2, padx=1, pady=1, sticky='nw')

        ##<~ upload button
        self.btn_upload = Button(self.fm1, width=10, text='UPLOAD', command=self.upload2server)
        self.btn_upload.grid(row=10, column=0, padx=4, pady=1, sticky='nw')
        self.progress_upload = Progressbar(self.fm1, orient=HORIZONTAL, length=100, mode='indeterminate') ##<~ progress bar for upload

        ##<~ send email
        self.btn_sendemail = Button(self.fm1, width=10, text='SEND Email', command=self.send_email)
        self.btn_sendemail.grid(row=11, column=0, padx=4, pady=1, sticky='nw')
        self.progress_sendemail = Progressbar(self.fm1, orient=HORIZONTAL, length=100, mode='indeterminate') ##<~ progress bar for sendemail

        # #<~ stop the process running in background
        self.root.protocol("WM_DELETE_WINDOW", self.close_app)

    def pick_ppt_report_file_path_dialog(self):
        self.e_ppt_report.delete(0, 'end')
        f_path = filedialog.askopenfilename(filetypes=(('Report file', '*.*'), ('all files', '*.*')))
        f_path = os.path.abspath(f_path)
        self.e_ppt_report.insert(INSERT, f_path)

    def pick_up2server_path_dialog(self):
        self.e_upload2server.delete(0, 'end')
        f_path = filedialog.askdirectory(initialdir=self.upload2server_fd_path)
        f_path = os.path.abspath(f_path)
        self.e_upload2server.insert(INSERT, f_path)

    def has_intranet(self, host, port, timeout=3):
        """intranect connectivity confirmation"""
        try:
            socket.setdefaulttimeout(timeout)
            socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect((host, port))
            return True
        except socket.error:
            return False

    def upload2server(self):
        """upload local folder to server"""
        def real_upload2server():
            self.progress_upload.grid(row=10, column=1, padx=1, pady=1, sticky='nw')
            self.progress_upload.start()

            ##<~ find out file name, which will be used as mail subject
            ppt_report_path = self.e_ppt_report.get()
            src_dir = os.path.split(ppt_report_path)[0]
            tmp_fd_name = os.path.split(src_dir)[-1]
            dst_dir = self.e_upload2server.get()
            dst_dir = os.path.join(dst_dir, tmp_fd_name)
            if self.has_intranet(self.server_address, self.port):
                copy_tree(src_dir, dst_dir)
                self.btn_upload.configure(bg='green') #<~ user-friendy
            else:
                messagebox.showinfo("Information", "Warning: Lost SSV intranet connection")

            self.progress_upload.stop()
            self.progress_upload.grid_forget()
            self.btn_upload['state'] = 'normal'
            self.btn_sendemail['state'] = 'normal'
        self.btn_upload['state'] = 'disabled'
        self.btn_sendemail['state'] = 'disabled'
        threading.Thread(target=real_upload2server).start()

    def send_email(self):
        """send email via outlook"""
        def real_send_email():
            self.progress_sendemail.grid(row=11, column=1, padx=1, pady=1, sticky=NW)
            self.progress_sendemail.start()

            ##<~ find out file name, which will be used as mail subject
            ppt_report_path = self.e_ppt_report.get()
            fn = os.path.split(ppt_report_path)[-1]
            ext = os.path.splitext(ppt_report_path)[-1]
            ##<~ attachment
            attachment = ppt_report_path #<~ local ppt file path
            ##<~ ssv server folder
            save2fd_path = os.path.split(ppt_report_path)[0]
            tmp_fd_name = os.path.split(save2fd_path)[-1]
            dst_dir = self.e_upload2server.get()
            dst_dir = os.path.join(dst_dir, tmp_fd_name)
            ##<~ send email
            report_mail = mail.UlMail(attachment)
            report_mail.send()
            self.btn_sendemail.configure(bg='green') #<~ user-friendy

            self.progress_sendemail.stop()
            self.progress_sendemail.grid_forget()
            self.btn_upload['state'] = 'normal'
            self.btn_sendemail['state'] = 'normal'
        self.btn_upload['state'] = 'disabled'
        self.btn_sendemail['state'] = 'disabled'
        threading.Thread(target=real_send_email).start()

    def close_app(self):
        if os.path.exists(self.tmp):
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
