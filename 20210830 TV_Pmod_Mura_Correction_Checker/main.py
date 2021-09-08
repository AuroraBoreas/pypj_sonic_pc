import sys, os
sys.path.append('.')

from library.excel.image import MuraImageLoader, logging, Path
from library.plrLog import SetLogConverter
from library.extractor import Extractor
import datetime

def clean_dst_history(dstBinFolder:Path, dstImageFolder:Path):
    for root, _, files in os.walk(dstBinFolder):
        for file in files:
            os.remove(os.path.join(root, file))
            
    for root, _, files in os.walk(dstImageFolder):
        for file in files:
            os.remove(os.path.join(root, file))

def main():
    # srcBinFolder   = r'\\43.98.232.61\电气\XU_J\AG65_MURA'
    pmod_name      = "AG65_ITC(CHN)"
    # dstBinFolder   = "data"
    # dstImageFolder = "Images"

    # clean_dst_history(dstBinFolder, dstImageFolder)
    # e1 = Extractor(srcBinFolder, dstBinFolder)
    # e1.copy()

    # plr_exe        = "PLRLog.exe"
    # dstBinFolder   = "data"
    # img_path       = "img.bmp"
    # dstImageFolder = "Images"
    
    # slc = SetLogConverter(plr_exe, dstBinFolder, img_path, dstImageFolder)
    # slc.to_image()

    template       = r'templates\Mura_List_Template.xlsx'
    dst_xl         = '{} {}_Mura_Image_List.xlsx'.format(datetime.date.today().strftime('%Y%m%d'), pmod_name)
    src_img_folder = 'Images'
    dstWidth      = 240
    dstCellName   = 'B'
    dstCellCol384 = 'C'
    dstCellCol640 = 'D'
    mil = MuraImageLoader(template, dst_xl, src_img_folder, dstWidth, dstCellName, dstCellCol384, dstCellCol640)
    mil.work()
    mil.clean()
    logging.info("Mura Images loaded into Excel successfully!")

if __name__ == '__main__':
    main()