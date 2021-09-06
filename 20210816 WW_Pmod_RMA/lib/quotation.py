"""
this module is to build a database and maintain price for each RMA pmod.

it has the following functionalities.
~ pickle maintained pmods
  >>> using non-sql or sql? using pandas? database?

~ read a given pmod-price list
  >>> this list is an excel format?

~ compare maintainable status
  >>> using l e m? or using in operator, time complexity = O(n), space complexity = O(n)

~ filter wait-to-maintain pmods
  >>> using pandas dataframe to filter, ie. df['bom_maintained'] = True;

~ transfer the wait-to-maintain pmods to "联络书（新品维护）ver2 210730.xlsx"
  >>> write the wait-to-maintain pmod list to it

~ change the status of the wait-to-maintain pmods to <maintained>
  >>> overwrite the status

~ add Warehouse code, MOQ, LT at the last line of the list
  >>> overwrite the comment

author
~ @ZL, 202010803

changelog
~ v0.01, initial build

"""

import sqlite3
import pandas as pd
import contextlib
import logging
from typing import (NewType)

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(message)s')

Path      = NewType('Path', str)
Cursor    = NewType('Cursor', sqlite3.Cursor)
DataFrame = NewType('DataFrame', pd.DataFrame)

class Tracker:
    def __init__(self, db:Path):
        self._db  = db

    @contextlib.contextmanager
    def __enter_table(self, cur:Cursor, act:str)->str:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS quotations(
                PMOD TEXT NOT NULL, 
                ModelName TEXT, 
                ProductNo TEXT, 
                Year INTEGER, 
                Month INTEGER, 
                Price FLOAT, 
                Unit TEXT,
                BuybackProductNo TEXT,
                BOM_Maintained BOOLEAN NOT NULL,
                BOM_MaintainDate TEXT,
                UNIQUE(PMOD, ModelName, ProductNo, Year, Month, Price, Unit, BuybackProductNo) ON CONFLICT IGNORE
            );
            """)
        try:
            yield
        finally:
            # cur.execute('DROP TABLE quotations;')
            logging.info(f'{act} completed')
            return f'{act} completed'
    
    def create(self, lcm_quotation:Path)->str:
        with sqlite3.connect(self._db) as conn:
            cur = conn.cursor()
            df  = pd.read_excel(lcm_quotation, sheet_name=0, header=0)
            info= 'create {} rows'.format(len(df))
            with self.__enter_table(cur, info):
                df.to_sql('quotations', conn, if_exists='append', index=False)
                # rv = pd.read_sql('SELECT * FROM quotations LIMIT 5', conn)
                # print(rv)
                return info 

    def __date_parser(self, d:DataFrame)->tuple:
        """
        src text: 適用開始日：2021年8月1日
        TODO: extract year, month
        """
        sep_date = r'適用開始日：'
        sep_year = r'年'
        sep_mon  = r'月'
        position = 1
        pos_date = -1
        pos_year = 0
        pos_mon  = 0

        date  = list(d.columns.values)[position]
        date  = date.split(sep_date)[pos_date]
        year  = date.split(sep_year)[pos_year]
        mon   = date.split(sep_year)[pos_date].split(sep_mon)[pos_mon]
        return (year, mon)

    def __unit_parser(self, u:DataFrame)->str:
        row, col = 1, 4
        return u.iloc[row, col]
        
    def emport(self, new_quotation:Path, is_bom_maintained:bool=0, bom_maintain_date:str='TBD')->None:
        sn = 0
        hd = 0
        sr = 14
        sf = 2
        df_du = pd.read_excel(new_quotation, sheet_name=sn, header=hd, skipfooter=sf, skiprows=sr)
        year, mon = self.__date_parser(df_du)
        unit = self.__unit_parser(df_du)
        # logging.info("year={}, month={}, unit={}".format(year, mon, unit)) # date, unit

        quotation_table_start_row = 3
        df:DataFrame = df_du[quotation_table_start_row:] # table
        df.columns = ['header0', 'header1', 'ModelName', 'ProductNo', 'Price', 'PMOD']
        df_clean = df[['PMOD','ModelName', 'ProductNo', 'Price']] # sequencing
        df_new = df_clean.assign(
          Year  = year,
          Month = mon,
          Unit  = unit,
          BuybackProductNo = df_clean['ProductNo'].apply(lambda x: x[1:] + x[0]),
          BOM_Maintained   = is_bom_maintained,
          BOM_MaintainDate = bom_maintain_date
        )
      
        with sqlite3.connect(self._db) as conn:
            cur  = conn.cursor()
            info = 'import {} rows'.format(len(df_new))
            with self.__enter_table(cur, info):
              df_new.to_sql('quotations', conn, if_exists='append', index=False)
              # rv = pd.read_sql('SELECT * FROM quotations LIMIT 5', conn)
              # logging.info(rv)

    def retrieve(self, is_bom_maintained:bool=0)->DataFrame:
        with sqlite3.connect(self._db) as conn:
            cur = conn.cursor()
            with self.__enter_table(cur, 'retrive'):
                df = pd.read_sql(
                  """
                  SELECT *
                  FROM quotations
                  WHERE BOM_Maintained={0}
                  """.format(is_bom_maintained), 
                conn)
                return df

    def update(self, old_bom_status:bool, new_bom_status:bool, bom_maintain_date:str)->None:
        with sqlite3.connect(self._db) as conn:
            cur = conn.cursor()
            with self.__enter_table(cur, 'update'):
                cur.execute(
                  """
                  UPDATE quotations
                  SET BOM_Maintained=(?), BOM_MaintainDate=(?)
                  WHERE BOM_Maintained=(?);
                  """, (new_bom_status, bom_maintain_date, old_bom_status)
                )

    def export(self, dst:Path)->None:
        with sqlite3.connect(self._db) as conn:
            cur = conn.cursor()
            with self.__enter_table(cur, 'export'):
              df = pd.read_sql(
                """
                SELECT *
                FROM quotations
                """,
              conn)
              df.to_excel(dst,index=False)
        
    def delete(self):
        pass
        
if __name__ == "__main__":
    db = 'test.db'
    qu = r'C:\Users\5106001995\Desktop\20210727 AG65_SET_LED_Mura\Sample BB\data\TENGO_Quotation_DB.xlsm'
    # nq = r'C:\Users\5106001995\Desktop\20210727 AG65_SET_LED_Mura\Sample BB\FY21 6TP LCM WW見積書  追加.xlsx'
    nq = r'C:\Users\5106001995\Desktop\20210727 AG65_SET_LED_Mura\Sample BB\data\FY21 7TP\FY21 8TP LCM WW見積書.xlsx'
    bt = Tracker(db)
    bt.create(qu)
    bt.emport(nq)