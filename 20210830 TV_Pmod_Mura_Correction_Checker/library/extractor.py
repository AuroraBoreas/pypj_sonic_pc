"""

this module is to extract some TV set MURA/CUC bin files from a given folder

@ZL, 20210603

"""

import pathlib
import shutil
import sys
sys.path.append('.')
from library.excel.image import logging

class Extractor:
    # ids: tuple = ("MesData128_", "MesData384_", "MesData640_", "MesData896_", "MesData1023_")
    ids:tuple = ("MesData896_", "MesData896Result_")

    def __init__(self, src_folder: str, to_folder: str):
        self.srcFolder = src_folder
        self.toFolder  = to_folder

    def __repr__(self):
        return f"\nSRC folder: {self.srcFolder}\nTo Folder: {self.toFolder}"

    def _indentify(self)->str:
        """
        patter: MesData128_, MesData384_, MesData640_, MesData896_, MesData1023_
        """
        files: list = sorted(pathlib.Path(self.srcFolder).rglob("*.bin"))
        for file in files:
            if file.name.startswith(self.ids):
                yield file

    def copy(self):
        for file in self._indentify():
            logging.info(file)
            shutil.copy(file, self.toFolder)

if __name__ == "__main__":
    e1: Extractor = Extractor(r"\\43.98.232.61\电气\Shen Yu\DATA\SOMA", ".")
    e1.copy()