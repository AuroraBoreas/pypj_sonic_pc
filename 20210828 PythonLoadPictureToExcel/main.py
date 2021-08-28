import sys
sys.path.append('.')

from library.excel.image import MuraImageLoader

if __name__ == '__main__':
    template       = r'templates\out.xlsx'
    dst_xl         = 'test1.xlsx'
    src_img_folder = r'data'
    dstWidth       = 200
    dstCellCol375  = 'B'
    dstCellCol675  = 'C'
    mil = MuraImageLoader(template, dst_xl, src_img_folder, dstWidth, dstCellCol375, dstCellCol675)
    mil.work()
    mil.clean()