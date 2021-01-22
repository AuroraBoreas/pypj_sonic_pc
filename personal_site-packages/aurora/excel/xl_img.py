import os, sys
import win32com.client as win32
from PIL import ImageGrab
from ctypes import windll

class Excel_Image_Extractor:
    def __init__(self, wb_path: str, ws_name: str, target_address: str, target_name: str, tmp_img_folder: str):
        self.wb_path     = wb_path #<~ must be a full path
        self.ws_name     = ws_name
        self.rng_address = target_address
        self.rng_name    = target_name
        self.img_folder  = tmp_img_folder

        self.xlapp               = win32.DispatchEx('Excel.Application') #NOT work if Excel application is not registered
        self.xlapp.DisplayAlerts = False
        self.xlapp.Visible       = False

    def __clear_clipboard(self):
        ##<~~clear clipboard and quit XL application
        if windll.user32.OpenClipboard(None):
            windll.user32.EmptyClipboard()
            windll.user32.CloseClipboard()

    def start(self, screenshotonly=False):
        wb = self.xlapp.Workbooks.Open(self.wb_path)
        ws = wb.Worksheets(self.ws_name)
        try:
            rng = ws.Range(self.rng_address)
            rng.Copy()
            img = ImageGrab.grabclipboard()
            img.save(os.path.join(self.img_folder, "{}.jpg".format(self.rng_name)), 'jpeg')

            if not screenshotonly:
                for _, shape in enumerate(ws.Shapes):
                    if shape.Name.startswith('Graph'):
                        shape.Copy()
                        img = ImageGrab.grabclipboard()
                        img.save(os.path.join(self.img_folder, "{}.jpg".format(shape.Name)), 'jpeg')

        except AttributeError:
            sys.exit("Failed to grab img from XL file {}".format(wb_path))
        finally:
            self.__clear_clipboard()
            wb.Close(False)
        return

    def quit(self):
        self.xlapp.Quit()

if __name__ == '__main__':
    # wb_path = "01.Brightness & chromaticity_jis_NX65.xlsm"
    wb_path = r"C:\Users\5106001995\Desktop\20201222 NX75_QC_Optical_NG\00_Brightness & chromaticity_NX75_MP.xlsm"
    tmp_fd  = r"C:\Users\5106001995\Desktop\temp"
    xi = Excel_Image_Extractor(wb_path, "report", "C4:K13", "summary", tmp_fd)
    xi.start()
    xi.quit()
