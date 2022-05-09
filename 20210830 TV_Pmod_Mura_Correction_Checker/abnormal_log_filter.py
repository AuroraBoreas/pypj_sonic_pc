import pathlib
import shutil
from typing import Optional

def sift_target_logs(src:str, dst:str, targets:list)->Optional[None]:
    files:list = sorted(pathlib.Path(src).rglob("*.bin"))
    ser:str = ""
    for file in files:
        ser = file.name.split('_')[2]
        # print(ser)
        if ser in targets:
            # print(file)
            shutil.copy(file, dst)


def main()->None:
    src:str = "data"
    dst:str = "FW75_Whiteline_log"
    tgt:list = [
        "0000020",
        "0000032",
        "0000049",
        "0000051",
        "0000052",
        "0000107",
        "0000108",
        "0000109",
        "0000114",
        "0000136",
        "0000241",
        "0000268",
        "0000269",
    ]
    sift_target_logs(src, dst, tgt)

if __name__ == '__main__':
    main()
