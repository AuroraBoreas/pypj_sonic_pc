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

"""

import openpyxl, os, contextlib
from PIL import Image
from typing import NewType, List
Path = NewType('Path', str)

@contextlib.contextmanager
def open_workbook(wb_path:Path, dst_path:Path):
    wb = openpyxl.load_workbook(wb_path)
    try:
        yield wb
    finally:
        wb.save(dst_path)
        wb.close()

class MuraImageLoader:
    img_ext     = '.jpg'
    muradata375 = '375'
    muradata675 = '675'
    dstStartRow = 2

    def __init__(self, xl_template:Path, dst_xl:Path, img_folder:Path, width:int, dstCellCol375:str, dstCellCol675:str):
        self._xl_template    = xl_template
        self._dst_xl         = dst_xl
        self._img_folder     = img_folder
        self.width           = width
        self.dstCellCol375   = dstCellCol375
        self.dstCellCol675   = dstCellCol675
        self._375:List[Path] = []
        self._675:List[Path] = []

    def _sort_images(self)->None:
        for root, _, files in os.walk(self._img_folder):
            for file in files:
                if os.path.splitext(file)[-1] == self.img_ext:
                    filepath = os.path.join(root, file)
                    if self.muradata375 in file:
                        self._375.append(filepath)
                    if self.muradata675 in file:
                        self._675.append(filepath)

    def _export(self)->None:
        idx_w, idx_h = (0, 1)
        i = j = self.dstStartRow
        tmp_img1 = 'tmp1.jpg'
        tmp_img2 = 'tmp2.jpg'
        with open_workbook(self._xl_template, self._dst_xl) as wb:
            ws = wb.worksheets[0]
            if self._375:
                for image in self._375:
                    img = Image.open(image)
                    width_percent = (self.width/float(img.size[idx_w]))
                    hsize = int((float(img.size[idx_h])*float(width_percent)))
                    img = img.resize((self.width, hsize), Image.ANTIALIAS)
                    img.save(tmp_img1)
                    img = openpyxl.drawing.image.Image(tmp_img1)
                    dstCellAddress = f'{self.dstCellCol375}{i}'
                    ws.add_image(img, dstCellAddress)
                    i += 1

            if self._675:
                for ximage in self._675:
                    img = Image.open(ximage)
                    width_percent = (self.width/float(img.size[idx_w]))
                    hsize = int((float(img.size[idx_h])*float(width_percent)))
                    img = img.resize((self.width, hsize), Image.ANTIALIAS)
                    img.save(tmp_img2)
                    img = openpyxl.drawing.image.Image(tmp_img2)
                    dstCellAddress = f'{self.dstCellCol675}{j}'
                    ws.add_image(img, dstCellAddress)
                    j += 1

    def clean(self)->None:
        self._375.clear()
        self._675.clear()

    def work(self)->None:
        self._sort_images()
        self._export()

if __name__ == '__main__':
    template       = r'C:\Users\Aurora_Boreas\Desktop\20210828 PythonLoadPictureToExcel\templates\out.xlsx'
    dst_xl         = r'C:\Users\Aurora_Boreas\Desktop\20210828 PythonLoadPictureToExcel\test.xlsx'
    src_img_folder = r'C:\Users\Aurora_Boreas\Desktop\20210828 PythonLoadPictureToExcel\data'
    dstWidth      = 200
    dstCellCol375 = 'B'
    dstCellCol675 = 'C'
    mil = MuraImageLoader(template, dst_xl, src_img_folder, dstWidth, dstCellCol375, dstCellCol675)
    mil.work()
    mil.clean()