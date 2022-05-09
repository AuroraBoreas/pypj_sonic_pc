import sys, os
sys.path.append('.')

from library.config import srcBinFolder
from library.excel.image import MuraImageLoader, logging, Path
from library.plrLog import SetLogConverter
from library.extractor import Extractor
import datetime

def clean_dst_history(dstBinFolder:Path, dstImgFolder:Path):
    for root, _, files in os.walk(dstBinFolder):
        for file in files:
            os.remove(os.path.join(root, file))
            
    for root, _, files in os.walk(dstImgFolder):
        for file in files:
            os.remove(os.path.join(root, file))

class MuraChecker:
    @staticmethod
    def clean_hist(dstBinFolder, dstImgFolder):
        clean_dst_history(dstBinFolder, dstImgFolder)

    @staticmethod
    def extract_src(srcBinFolder, dstBinFolder, interval=1, ire=896)->None:
        e = Extractor(srcBinFolder, dstBinFolder)
        e.ids_before = (f"MesData{ire}_", )
        e.ids_after  = (f"MesData{ire}Result_", )
        e.copy(interval)

    @staticmethod
    def convert_img(plr_path, dstBinFolder, img_path, dstImgFolder)->None:
        slc = SetLogConverter(plr_path, dstBinFolder, img_path, dstImgFolder)
        slc.to_image()

    @staticmethod
    def make_report(template, dst_xl, src_img_folder, dstWidth, dstCellName, dstCellColO, dstCellColR, ire)->MuraImageLoader:
        m = MuraImageLoader(template, dst_xl, src_img_folder, dstWidth, dstCellName, dstCellColO, dstCellColR)
        m.idn_muradata = (f"MesData{ire}_", )
        m.idn_muradataR = (f"MesData{ire}Result_", )
        return m

def main():
    pmod_name      = "FW75"
    dstBinFolder   = "data"
    dstImgFolder   = "Images"
    plr_exe        = "PLRLog.exe"
    img_path       = "img.bmp"
    ire            = 640 # 128, 384, 640, 896, 1024
    template       = r'templates\Mura_List_Template.xlsx'
    dst_xl         = '{}_{}_Mura_Image_List_{}.xlsx'.format(datetime.datetime.today().strftime('%Y%m%d%H%M%S'), pmod_name, ire)
    src_img_folder = 'Images'
    dstWidth       = 240
    dstCellName    = 'B'
    dstCellColO    = 'C'
    dstCellColR    = 'D'
    interval       = 1

    logging.info("starting..")
    mura_checker = MuraChecker()
    # mura_checker.clean_hist(dstBinFolder, dstImgFolder)
    # mura_checker.extract_src(srcBinFolder, dstBinFolder, interval, ire)
    mura_checker.convert_img(plr_exe, dstBinFolder, img_path, dstImgFolder)
    m = mura_checker.make_report(template, dst_xl, src_img_folder, dstWidth, dstCellName, dstCellColO, dstCellColR, ire)
    m.work(m.cmp_both)
    logging.info('Done')

if __name__ == '__main__':
    main()