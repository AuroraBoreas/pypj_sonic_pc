import pathlib, os
import win32com.client as win32

class ReportMail:
    """
    this module is specialization for sending PowerPoint Report via email

    it has the following functionalities:
    - control Outlook Application
    - construct email body in HTML format
    - send email
    - quit Outlook Application

    about
    - author: @ZL, 20181223

    """
    def __init__(self, ppt_name: str, dst_dir: str, slide_img_folder: str, src_ppt_file: str, reciever = "Liang.Zhang@sony.com"):
        self.__outlook        = win32.Dispatch('outlook.application')
        self.__mail           = self.__outlook.CreateItem(0)
        self.ppt_name         = ppt_name
        self.dst_dir          = dst_dir
        self.slide_img_folder = slide_img_folder
        self.src_ppt_file     = src_ppt_file
        self.reciever         = reciever

    def __attachments(self):
        img_id: int         = 1
        img_body_slide: str = ""

        for file in sorted(pathlib.Path(self.slide_img_folder).glob("*.png")):
            # attachment
            fp: str    = str(file.absolute())
            attachment = self.__mail.Attachments.Add(fp) # abs path only
            # how: [https://stackoverflow.com/questions/44544369/i-am-not-able-to-add-an-image-in-email-body-using-python-i-am-able-to-add-a-pi]
            attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", f"MyId{img_id}")
            img_body_slide += f"""<p><img src="cid:MyId{img_id}"></p><br>"""
            img_id += 1

        return img_body_slide

    def __email_body(self):
        # HTML Body
        email_body_txt: str = """
        <p><font face="Arial">Hello all,</font></p>
        <p><font face="Arial">Pls refer to the attachment or the following hyperlink to see details of {0} result.<br>
        Server: <a href="{1}">SSV Server link</a></font></p>
        """.format(self.ppt_name, self.dst_dir)
        
        return email_body_txt

    def send(self):
        # Email headers
        self.__mail.To      = self.reciever
        self.__mail.Subject = self.ppt_name
        # HTML Body
        email_body_txt: str = self.__email_body()
        # report slides
        email_body_txt += self.__attachments()
        email_body_txt += """<p><font face="Arial">Regards,</font></p>"""
        # attach ppt file
        self.__mail.Attachments.Add(self.src_ppt_file) # abs path only
        self.__mail.HTMLBody = email_body_txt
        self.__mail.Send()

    def clean(self):
        for file in sorted(pathlib.Path(self.slide_img_folder).glob("*.png")):
            fp: str = str(file.absolute())
            os.remove(fp)

    def quit(self):
        self.clean()
        # self.outlook.Quit() # its not necessary to close outlook app. it askes for trouble