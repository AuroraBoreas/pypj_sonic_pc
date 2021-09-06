"""

this module is specialization for sending PowerPoint Report via email

it contains the following functionalities:
- control PowerPoint Application
- view specific slides in order
- shoot image for each slide
- quit PowerPoint Application

about
- author: @ZL, 20201223

"""

import win32com.client as win32
import time, sys, os
sys.path.append('.')
from library.shoot_image import shoot_img

class ReportPowerPoint:
    __start_slide_no = 2
    __ulimit = 5
    __wait   = .5
    __ext    = '.pptx'
    
    # PPT, Zoom leevl; 70%
    x1, y1 = 443, 276
    x2, y2 = 1781, 1033

    def __init__(self, strFile):
        self.strFile             = strFile
        self.__app               = win32.Dispatch('PowerPoint.application')
        self.__app.DisplayAlerts = False
        self.__app.Visible       = True
        self.__prs               = self.__app.Presentations.Open(strFile) # strFile MUST be an absolute path

    def quit(self):
        self.__app.Quit()

    def ttl_slides(self):
        return self.__prs.Slides.Count

    def shoot(self):
        if os.path.splitext(self.strFile)[-1] == self.__ext:
            if self.ttl_slides() < self.__ulimit:
                for i in range(self.__start_slide_no, self.ttl_slides()):
                    self.__prs.Slides(i).Select()
                    time.sleep(self.__wait)
                    self.__prs.Slides(i).Select()
                    shoot_img(f"images\\slide{i}.png", self.x1, self.y1, self.x2, self.y2)
            else:
                # summary page only
                self.__prs.Slides(self.__start_slide_no).Select()
                time.sleep(self.__wait)
                self.__prs.Slides(self.__start_slide_no).Select()
                shoot_img(f"images\\slide{self.__start_slide_no}.png", self.x1, self.y1, self.x2, self.y2) 

if __name__ == "__main__":
    ppt_file = r"C:\Users\5106001995\Desktop\20210517 AG65_MP_LCM(A1)\20210517 AG65_MP_LCM(A1).pptx"
    ppt = ReportPowerPoint(ppt_file)
    ppt.shoot()
    ppt.quit()
