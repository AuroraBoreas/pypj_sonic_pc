"""

this module is to extract some TV set MURA/CUC bin files from a given folder

@ZL, 20210603

"""

import pathlib
import shutil
import sys
import re
sys.path.append('.')
from library.excel.image import logging

class Extractor:
    # ids: tuple = ("MesData128_", "MesData384_", "MesData640_", "MesData896_", "MesData1023_")
    ids_before:tuple = ("MesData896_", "MesData2_")
    ids_after:tuple  = ("MesData896Result_", "MesData2Result_")

    def __init__(self, src_folder: str, to_folder: str):
        self.srcFolder = src_folder
        self.toFolder  = to_folder

    def __repr__(self):
        return f"\nSRC folder: {self.srcFolder}\nTo Folder: {self.toFolder}"

    def _indentify(self, interval)->str:
        """
        pattern: MesData128_, MesData384_, MesData640_, MesData896_, MesData1023_
        """
        files: list = sorted(pathlib.Path(self.srcFolder).rglob("*.bin"))
        rfiles:list = reversed(files)
        rv_before:dict = {}
        rv_after:dict = {}
        fn:str = ""
        i:int = 0; j:int = 0
        for file in rfiles:
            fn = re.sub(r"_[0-9]{14}", "", file.name)
            if file.name.startswith(self.ids_before):
                if fn not in rv_before:
                    rv_before[fn] = file
            if file.name.startswith(self.ids_after):
                if fn not in rv_after:
                    rv_after[fn] = file
                
        # print(len(rv_before.values()))
        for val in rv_before.values():
            if i % interval == 0:
                # logging.info(val)
                yield val
            i += 1
        # print(len(rv_after.values()))
        for val in rv_after.values():
            if j % interval == 0:
                # logging.info(val)
                yield val
            j += 1

    def copy(self, interval=1):
        for file in self._indentify(interval):
            # logging.info(file)
            shutil.copy(file, self.toFolder)


if __name__ == "__main__":
    e1: Extractor = Extractor(r"\\43.98.232.61\电气\trashbox\Shen Yu\DATA", ".")
    e1.copy(20)