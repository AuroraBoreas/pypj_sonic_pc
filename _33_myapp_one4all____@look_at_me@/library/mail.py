"""
this module is specialization for sending PowerPoint Report via email

it contains the following functionalities:
- control Outlook Application
- construct email body in HTML format
- send email
- quit Outlook Application

about
- author: @ZL, 20201223

"""

import win32com.client as win32
import pathlib
import os

class ReportMail:
    def __init__(self, ppt_name, dst_dir, slide_img_folder, src_ppt_file):
        self.__outlook          = win32.Dispatch('outlook.application')
        self.__mail             = self.__outlook.CreateItem(0)
        self.ppt_name           = ppt_name
        self.dst_dir            = dst_dir
        self.slide_img_folder   = slide_img_folder
        self.src_ppt_file       = src_ppt_file

    def attachments(self):
        img_id: int         = 1
        img_body_slide: str = ""
        for file in sorted(pathlib.Path(self.slide_img_folder).glob("*.png")):
            ##<~ attachment
            fp: str    = str(file.absolute())
            attachment = self.__mail.Attachments.Add(fp) # abs path only
            ## how: [https://stackoverflow.com/questions/44544369/i-am-not-able-to-add-an-image-in-email-body-using-python-i-am-able-to-add-a-pi]
            attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", f"MyId{img_id}")
            img_body_slide += f"""<p><img src="cid:MyId{img_id}"></p><br>"""
            img_id += 1
        return img_body_slide

    def send(self):
        #<~ Email headers
        self.__mail.To      = 'Liang.Zhang@sony.com' #<~ this is safer
        self.__mail.Subject = f'{self.ppt_name}'
        #<~HTML Body
        email_body_txt: str = """
        <p><font face="Arial">Hello all,</font></p>
        <p><font face="Arial">Pls refer to the attachment or the following hyperlink to see details of <b>{0}</b> result.<br>
        Server: <a href="{1}">SSVE Server link</a></font></p>
        """.format(self.ppt_name, self.dst_dir)
        #<~report slides
        email_body_txt += self.attachments()
        email_body_txt += """<p><font face="Arial">Regards,</font></p>"""
        #<~attach ppt file
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

if __name__ == "__main__":
    rpt_mail = ReportMail("NX75", '.', '.', r"C:\Users\5106001995\Desktop\_102_ppt_to_mail\20201222 NX75_MP_Optical_Evaluation.pptx")
    rpt_mail.send()
    rpt_mail.quit()
