from library.checker import (datetime, logging, MuraChecker, srcBinFolder)

def main():
    pmod_name      = "FW75" # pmod name
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
    interval       = 1 # sampling

    logging.info("starting..")
    mura_checker = MuraChecker()
    mura_checker.clean_hist(dstBinFolder, dstImgFolder)
    mura_checker.extract_src(srcBinFolder, dstBinFolder, interval, ire)
    mura_checker.convert_img(plr_exe, dstBinFolder, img_path, dstImgFolder)
    m = mura_checker.make_report(template, dst_xl, src_img_folder, dstWidth, dstCellName, dstCellColO, dstCellColR, ire)
    m.work(m.cmp_both)
    logging.info('Done')

if __name__ == '__main__':
    main()