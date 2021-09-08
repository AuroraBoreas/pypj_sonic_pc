"""
this module loads Mura Images into an excel template;

it has the following functionalities
- sort images from a given source image folder
- open template excel workbook
- load resized images into specific columns in the workbook

author
- @ZL, 20210828

changelog
- v0.01, initial build

Ref links:
- https://stackoverflow.com/questions/51591634/fill-a-image-in-a-cell-excel-using-python
- https://stackoverflow.com/questions/2232742/does-python-pil-resize-maintain-the-aspect-ratio
- https://stackoverflow.com/questions/1405602/how-to-adjust-the-quality-of-a-resized-image-in-python-imaging-library

"""

import openpyxl, os, contextlib
from PIL import Image
from typing import NewType, List
Path = NewType('Path', str)
import logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(message)s')

@contextlib.contextmanager
def open_workbook(wb_path:Path, dst_path:Path):
    wb = openpyxl.load_workbook(wb_path)
    try:
        yield wb
    finally:
        wb.save(dst_path)
        wb.close()

class MuraImageLoader:
    img_ext     = '.bmp'
    muradata896 = 'mesdata896_'
    muradataresult896 = 'mesdata896result_'
    dstStartRow = 2
    img_quality = 95

    def __init__(self, xl_template:Path, dst_xl:Path, img_folder:Path, width:int, dstCellName:str, dstCellCol384:str, dstCellCol640:str):
        self._xl_template     = xl_template
        self._dst_xl          = dst_xl
        self._img_folder      = img_folder
        self.width            = width
        self.dstCellName      = dstCellName
        self.dstCellCol384    = dstCellCol384
        self.dstCellCol640    = dstCellCol640
        self._896:List[Path]  = []
        self._896R:List[Path] = []

    def _sort_images(self)->None:
        for root, _, files in os.walk(self._img_folder):
            for file in files:
                if file.endswith(self.img_ext):
                    filepath = os.path.join(root, file)
                    # logging.info(filepath)
                    if self.muradata896 in file.lower():
                        self._896.append(filepath)
                    if self.muradataresult896 in file.lower():
                        self._896R.append(filepath)
        if not(len(self._896) == len(self._896R) and len(self._896) > 0 and len(self._896R) > 0):
            raise ValueError("Before/After QTY::SRC_Images Not Match, _896 = {}, _896Result = {}".format(len(self._896), len(self._896R)))

    def _is_pmod_id_match(self, x:str, y:str)->bool:
        delimiter = '_'
        idx_id    = 2
        return x.split(delimiter)[idx_id] == y.split(delimiter)[idx_id]

    def _export(self)->None:
        idx_w, idx_h = (0, 1)
        i = j = self.dstStartRow
        tmp_img1 = 'tmp1.jpg'
        tmp_img2 = 'tmp2.jpg'

        with open_workbook(self._xl_template, self._dst_xl) as wb:
            ws = wb.worksheets[0]
            ub = len(self._896)
            for k in range(0, ub):
                img_896 = self._896[k]
                img_896R = self._896R[k]
                if not self._is_pmod_id_match(img_896, img_896R):
                    raise FileNotFoundError(f"Pair {self.muradata896}:{self.muradataresult896}, PmodID Not Match by")
                
                img = Image.open(img_896)
                width_percent = (self.width/float(img.size[idx_w]))
                hsize = int((float(img.size[idx_h])*float(width_percent)))
                img = img.resize((self.width, hsize), Image.ANTIALIAS)
                img.save(tmp_img1, quality=self.img_quality)
                img = openpyxl.drawing.image.Image(tmp_img1)
                dstCellAddress = f'{self.dstCellName}{i}'
                ws[dstCellAddress].value = os.path.split(img_896)[-1]
                dstCellAddress = f'{self.dstCellCol384}{i}'
                ws.add_image(img, dstCellAddress)
                i += 1

                img = Image.open(img_896R)
                width_percent = (self.width/float(img.size[idx_w]))
                hsize = int((float(img.size[idx_h])*float(width_percent)))
                img = img.resize((self.width, hsize), Image.ANTIALIAS)
                img.save(tmp_img2, quality=self.img_quality)
                img = openpyxl.drawing.image.Image(tmp_img2)
                # dstCellAddress = f'{self.dstCellName}{i}'
                # ws[dstCellAddress].value = os.path.split(image)[-1]
                dstCellAddress = f'{self.dstCellCol640}{j}'
                ws.add_image(img, dstCellAddress)
                j += 1

    def clean(self)->None:
        self._896.clear()
        self._896R.clear()

    def work(self)->None:
        self._sort_images()
        # print("Before Adj 896")
        # print(self._896)
        # print("After Adj 896")
        # print(self._896R)
        self._export()

if __name__ == '__main__':
    template       = r'C:\Users\Aurora_Boreas\Desktop\pypj_sonic_pc\20210828 PythonLoadPictureToExcel\templates\out.xlsx'
    dst_xl         = r'C:\Users\Aurora_Boreas\Desktop\pypj_sonic_pc\20210828 PythonLoadPictureToExcel\test1.xlsx'
    src_img_folder = r'C:\Users\Aurora_Boreas\Desktop\pypj_sonic_pc\20210828 PythonLoadPictureToExcel\data'
    dstWidth      = 240
    dstCellName   = 'B'
    dstCellCol384 = 'C'
    dstCellCol640 = 'D'
    mil = MuraImageLoader(template, dst_xl, src_img_folder, dstWidth, dstCellCol384, dstCellCol640)
    mil.work()
    mil.clean()