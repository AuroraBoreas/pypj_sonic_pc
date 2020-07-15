"""
============================================================================================================================
#Z.Liang, 20191031
#v0, make life easier: to upload to server and send email automatically.
#v1, fix a potential server drive path problem
============================================================================================================================
"""

import os, sys, tempfile, threading, socket, distutils.dir_util
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

class App_win():
    favicon_path = r"data\assets\favicon_diamond.png"
    upload2server_fd_path = r"\\43.98.1.18\ssv sharefile\SHES-C\03-Info-Share\03-FYxx Model\FY20_Model\SSV Design model\NX\31.MP"
    server_address = '43.98.1.18'
    port = 445

    ##<~ tmp folder here is NOT for this module. it's for integral GUI<all4one>.
    tmp = os.path.join(tempfile.gettempdir(), '.{}'.format(hash(os.times())))
    if not os.path.isdir(tmp): os.makedirs(tmp)

    def __init__(self):
        """===GUI==="""
        self.root = Toplevel()  ##<~ toplevel when import from other module
        self.root.title("ToServer & E-Mail v.0, by Z.Liang, 20191031")

        self.w, self.h = 760, 140
        self.ws = self.root.winfo_screenwidth()
        self.hs = self.root.winfo_screenheight()
        x = int((self.ws - self.w) / 2)
        y = int((self.hs - self.h) / 2)
        self.root.geometry("{}x{}+{}+{}".format(self.w, self.h, x, y))

        try:
            self.imgicon = PhotoImage(file=self.favicon_path) ##<~ favicon setting up
            self.root.tk.call('wm', 'iconphoto', self.root._w, self.imgicon)    ##<~ favicon setting up
        except tkinter._tkinter.TclError:
            pass
        self.root.focus_set()

        ##<~frame1
        self.fm1 = Frame(self.root)
        self.fm1.grid(row=0, column=0, padx=2, sticky='nw')

        self.lbl_ppt_report = Label(self.fm1, text='Report file:')
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
        except socket.error as ex:
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

            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)

            ##<~ find out file name, which will be used as mail subject
            ppt_report_path = self.e_ppt_report.get()
            fn = os.path.split(ppt_report_path)[-1]
            subject = fn.partition('.')[0]

            mail.To = 'Liang.Zhang@sony.com' # this is safer
            mail.Subject = subject
            ##<~ attachment
            attachment = ppt_report_path #<~ local ppt file path
            ##<~ ssv server folder
            save2fd_path = os.path.split(ppt_report_path)[0]
            tmp_fd_name = os.path.split(save2fd_path)[-1]
            dst_dir = self.e_upload2server.get()
            dst_dir = os.path.join(dst_dir, tmp_fd_name)
            #<~HTML Body
            email_body_txt = """
            <p><font face="Arial">Hello all, </font></p>
            <p><font face="Arial">Pls refer to the attachment or the following hyperlink to see details of {0}.<br>
            Server: <a href="{1}">SSV Server link</a></font></p>
            <p><font face="Arial">Regards,</font></p>
            """.format(subject, dst_dir)
            mail.HTMLBody = email_body_txt
            mail.Attachments.Add(attachment)
            mail.Send()
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
