import pyautogui, time
import win32com.client as win32

def shoot_img(str_img_name, x1, y1, x2, y2):
    """
    # shoot image for specific area
    take screenshot -> convert to binary data -> put on clipboard
    """
    w = x2 - x1
    h = y2 - y1
    pyautogui.screenshot(str_img_name, region=(x1, y1, w, h))

class ReportPowerPoint:
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
    __start_slide_no = 2

    def __init__(self, strFile: str, slide_img_folder: str):
        self.__app               = win32.Dispatch('PowerPoint.application')
        self.__app.DisplayAlerts = False
        self.__app.Visible       = True
        self.__prs               = self.__app.Presentations.Open(strFile) # strFile MUST be an absolute path
        self.slide_img_folder    = slide_img_folder

    def quit(self):
        self.__app.Quit()

    def ttl_slides(self):
        return self.__prs.Slides.Count

    def shoot(self):
        for i in range(self.__start_slide_no, self.ttl_slides()):
            self.__prs.Slides(i).Select()
            time.sleep(.5)
            self.__prs.Slides(i).Select()
            shoot_img(f"{self.slide_img_folder}\\slide{i}.png", 444, 278, 1779, 1031) # depends on personal PC and PowerPoint application setting, YMMV
