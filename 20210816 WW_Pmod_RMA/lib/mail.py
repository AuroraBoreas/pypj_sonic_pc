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
from lib.quotation import logging

from typing import NewType
Html = NewType('Html', str)
Path = NewType('Path', str)

class Mailer:
    _rev = 'Liang.Zhang@sony.com' #<~ this is safer

    def __init__(self, subject:str, body:Html, src_file:Path):
        self.__outlook = win32.Dispatch('outlook.application')
        self.__mail    = self.__outlook.CreateItem(0)
        self.subject   = subject
        self.body      = body
        self.src_file  = src_file
        
    def send(self):
        #<~ Email headers
        self.__mail.To      = self._rev
        self.__mail.Subject = self.subject
        #<~HTML Body
        email_body_txt      = self.body
        #<~report slides
        email_body_txt += """<p><font face="Arial">Regards,</font></p>"""
        #<~attach ppt file
        self.__mail.Attachments.Add(self.src_file) # abs path only
        self.__mail.HTMLBody = email_body_txt
        self.__mail.Send()
        logging.info('sent email successed!')