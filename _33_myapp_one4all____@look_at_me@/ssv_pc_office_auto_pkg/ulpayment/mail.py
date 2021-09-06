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
import os

class UlMail:
    def __init__(self, src_file):
        self.__outlook = win32.Dispatch('outlook.application')
        self.__mail    = self.__outlook.CreateItem(0)
        self.src_file  = src_file
        
    def send(self):
        #<~ Email headers
        self.__mail.To      = 'Liang.Zhang@sony.com' #<~ this is safer
        self.__mail.Subject = 'UL GmbH跟踪服务费支付依赖'
        payment_no = os.path.split(self.src_file)[-1].split('.')[0]
        #<~HTML Body
        email_body_txt: str = """
        <p><font face="DengXian">杨SAN, 金SAN
        <br>两位好。
        <br>
        <br>UL GmbH跟踪服务费支付依赖[编号 <u>{0}</u>]<b>已经申请</b>。系统审批<b>流转中</b>。
        <br>后续请协助<b>处理支付</b>。非常感谢。
        <br></font></p>
        """.format(payment_no)
        #<~report slides
        email_body_txt += """<p><font face="Arial">Regards,</font></p>"""
        #<~attach ppt file
        self.__mail.Attachments.Add(self.src_file) # abs path only
        self.__mail.HTMLBody = email_body_txt
        self.__mail.Send()

if __name__ == "__main__":
    rpt_mail = UlMail(r"C:\Users\5106001995\Desktop\20210518 ULGmbH_payment\SNX03C-2105-00179.htm")
    rpt_mail.send()
