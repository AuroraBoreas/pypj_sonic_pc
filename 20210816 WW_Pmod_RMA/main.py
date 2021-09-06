from lib.quotation import Tracker, logging
from lib.bom import generate_bom_excel
import os, datetime

if __name__ == "__main__":
    # db = r'C:\Users\5106001995\Desktop\20210727 AG65_SET_LED_Mura\Sample BB\lcm_ww_quotations.db'
    db = r'C:\Users\5106001995\Desktop\20210727 AG65_SET_LED_Mura\Sample BB\test.db'
    bs = r'C:\Users\5106001995\Desktop\20210727 AG65_SET_LED_Mura\Sample BB\data\TENGO_Quotation_DB.xlsm'
    folder = r'C:\Users\5106001995\Desktop\20210727 AG65_SET_LED_Mura\Sample BB\data\FY21 7TP'

    bt = Tracker(db)
    
    # @TODO: create table based on an existing excel workbook
    # bt.create(bs)

    # @TODO: export to excel
    # exported = 'exported.xlsx'
    # bt.export(exported)

    # @TODO: append data from new lcm_ww_quotation files into sqlite database recursively
    # for root, _, files in os.walk(folder):
    #     for file in files:
    #         bt.emport(os.path.join(root, file), is_bom_maintained=1, bom_maintain_date='2021-08-03 00:00:00')
    
    # @TODO: query wait-to-register quotations
    # rv = bt.retrieve(2021, 8, is_bom_maintained=0)

    # @TODO: write them into bom-registeration-excel
    # df = rv[['PMOD', 'BuybackProductNo', 'Price', 'ModelName']]
    # df_bom = df.assign(
    #     状态=r'部品',
    #     维护到=r'ASSY'
    # )
    
    # if len(df_bom) > 0:
    #     tmp_bom = r'C:\Users\5106001995\Desktop\20210727 AG65_SET_LED_Mura\Sample BB\data\templates\bom\联络书（新品维护）ver2 210730.xlsx'
    #     today   = datetime.datetime.today().strftime("%Y%m%d")
    #     dst     = '联络书（新品维护）ver2 210730_{0}.xlsx'.format(today)
    #     df_xls  = df_bom.iloc[0:9]
    #     # print(df_xls)
    #     generate_bom_excel(tmp_bom, df_xls, dst)
    # else:
    #     logging.info('all BuybackProductNo registered already!')
    
    # @TODO: update bom maintained status
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    bt.update(0, 1, now)
    
    # @TODO: send the bom-registeration-excel as an attachement within an email request
