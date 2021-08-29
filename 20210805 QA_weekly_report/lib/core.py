"""
this module reads all WW SMLD src file and aggregates into one file.

it has the following functionalities.
- enumerate all source files from a given folder
- clean defects dataframe
- format inputs dataframe
- dump cleaned dataframes into sqlite3 database

Author
- @ZL, 20210804

Changelog
- v0.01, initial build
- v0.02, clean and format algorithms
- v0.03, change design pattern: split into two dataframes: defects, inputs
- v0.04, add more columns to prepare for visualization
- v0.05, fix TVPlant column, "SO'EM" -> "SOEM"

"""

import datetime, time, os, logging, sqlite3, sys, contextlib, functools
sys.path.append('.')
import pandas as pd
from lib.date import utils
from typing import NewType, Dict, Any, Union, List, Callable
from lib import setname_mapping

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(message)s')
Path      = NewType('Path', str)
DataFrame = NewType('DataFrame', pd.DataFrame)
Cursor    = NewType('Cursor', sqlite3.Cursor)

def timer(func:Callable)->Callable:
    @functools.wraps(func)
    def inner(*args:Any, **kwargs:Any)->Any:
        beg = time.perf_counter()
        rv  = func(*args, **kwargs)
        end = time.perf_counter()
        logging.info('Successed. time lapsed(s): {0:.2f}'.format(end - beg))
        return rv
    return inner

class Merger:
    def __init__(self, folder:Path, db:Path,
                fix_defects_setname:str=None, fix_defects_setname_mapping:Dict[Any, Any]=None,
                fix_inputs_modelname:str=None, fix_inputs_modelname_mapping:Dict[Any,Any]=None,
                fix_inputs_fy:str=None, fix_inputs_fy_mapping:Dict[Any,Any]=None,
                fix_inputs_wkcode:str=None
    ):
        """read preprocessed source excel files(*.xlsx) from a given directory, and dump into database

        :param folder: a folder that holds source excel files(*.xlsx)
        :type folder: Path
        :param db: database
        :type db: Path
        :param fix_defects_setname: [description], defaults to None
        :type fix_defects_setname: str, optional
        :param fix_defects_setname_mapping: [description], defaults to None
        :type fix_defects_setname_mapping: Dict[Any, Any], optional
        :param fix_inputs_modelname: [description], defaults to None
        :type fix_inputs_modelname: str, optional
        :param fix_inputs_modelname_mapping: [description], defaults to None
        :type fix_inputs_modelname_mapping: Dict[Any,Any], optional
        :param fix_inputs_fy: [description], defaults to None
        :type fix_inputs_fy: str, optional
        :param fix_inputs_fy_mapping: [description], defaults to None
        :type fix_inputs_fy_mapping: Dict[Any,Any], optional
        :param fix_inputs_wkcode: [description], defaults to None
        :type fix_inputs_wkcode: str, optional
        """
        self._folder = folder
        self._db     = db

        self.fix_defects_setname          = fix_defects_setname
        self.fix_defects_setname_mapping  = fix_defects_setname_mapping
        self.fix_inputs_modelname         = fix_inputs_modelname
        self.fix_inputs_modelname_mapping = fix_inputs_modelname_mapping
        self.fix_inputs_fy                = fix_inputs_fy
        self.fix_inputs_fy_mapping        = fix_inputs_fy_mapping
        self.fix_inputs_wkcode            = fix_inputs_wkcode

        self._dfDefect:DataFrame = None
        self._dfInputs:DataFrame = None
    
    def _filter(self):
        """filter source files, especially excel temporary files which start with '~$'

        :yield: excel file
        :rtype: path-like string
        """
        xl_ext:str = '.xlsx'
        tmp_xl:str = '~$'
        idns:list = ['smld', 'quality']

        for root, _, files in os.walk(self._folder):
            for file in files:
                ext = os.path.splitext(file)[-1]
                if any(idn in file.lower() for idn in idns) and not file.startswith(tmp_xl) and xl_ext == ext:
                    logging.info(file)
                    yield os.path.abspath(os.path.join(root, file))

    @contextlib.contextmanager
    def __enter_defects_table(self, name:str, cur:Cursor, act:str)->None:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS {0}(
                No INTEGER,
                ConfirmationWC INTEGER,
                OccurredWC INTEGER,
                source TEXT,
                OccurredDate TEXT,
                ConfirmDate TEXT,
                TVPlant TEXT,
                LineName TEXT,
                Position TEXT,
                SetName TEXT,
                ModelName TEXT,
                ModuleID TEXT,
                WeekCode INTEGER,
                SonyPN TEXT,
                SITE TEXT,
                CellModel TEXT,
                CellID TEXT,
                NoticedSymptom TEXT,
                OSSorFFAjudgement TEXT,
                Classify TEXT,
                FY_mod TEXT,

                UNIQUE(ConfirmationWC, TVPlant, SetName, ModelName, ModuleID, Classify) ON CONFLICT IGNORE
            );
            """.format(name))
        try:
            yield
        finally:
            logging.info(f'SQL {act} completed')

    @contextlib.contextmanager
    def __enter_inputs_table(self, name:str, cur:Cursor, act:str)->None:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS {0}(
                TVPlant TEXT,
                ModelName TEXT,
                WeekCode INTEGER,
                QTY INTEGER,
                FY_mod TEXT,
                FY_MM TEXT,
                UNIQUE(TVPlant, ModelName, WeekCode, QTY) ON CONFLICT IGNORE
            );
            """.format(name))
        try:
            yield
        finally:
            logging.info(f'SQL {act} completed')

    def _match_fymod(self, x:str, d:Dict[Any,Any])->Union[str, None]:
        """return FY21 based on the unique character inside a give string x;

        :param x: pmod name. i.e: AG65
        :type x: str
        :param d: a dictionary contains unique character and FYxx pairs
        :type d: Dict[Any,Any]
        :return: string if found else None
        :rtype: Union[str, None]
        """
        for k in d.keys():
            if k in x:
                return d[k]
        else:
            return None
    
    def _clean_dfDefects(self, DataFrameList:List[DataFrame])->None:
        """clean defects DataFrame before dumping into database

        :param DataFrameList: a list contains raw dataframe
        :type DataFrameList: List[DataFrame]
        """
        # ~ Defects
        na_col:str           = 'Module ID'
        defects_cols:int     = 21
        defects_col_date:str = 'Confirm date'
        defects_na_date  = pd.Timestamp('20210101')
        defects_str_col1 = 'Occurred date'
        defects_str_col2 = 'Confirm date'
        defects_col_FYMM = 'FY_mod'

        _na = 'na'

        new_column_names   = {
            "No"                    :   "No",
            "Confirmation W/C"      :   "ConfirmationWC",
            "Occurred W/C"          :   "OccurredWC",
            "source"                :   "source",
            "Occurred date"         :   "OccurredDate",
            "Confirm date"          :   "ConfirmDate",
            "TV Plant"              :   "TVPlant",
            "LineName Line / SBI"   :   "LineName",
            "Position"              :   "Position",
            "Set name"              :   "SetName",
            "Model name"            :   "ModelName",
            "Module ID"             :   "ModuleID",
            "Week-Code"             :   "WeekCode",
            "Sony PN"               :   "SonyPN",
            "SITE"                  :   "SITE",
            "Cell Model"            :   "CellModel",
            "Cell-ID"               :   "CellID",
            "Noticed symptom"       :   "NoticedSymptom",
            "OSS or FFA judgement"  :   "OSSorFFAjudgement",
            "Classify"              :   "Classify",
            "FY_mod"                :   "FY_mod",
        }

        defects_tv_plants = 'TVPlant'
        defects_tv_plants_mapping = {"SO\'EM" : 'SOEM'}

        df:DataFrame = pd.concat(DataFrameList, ignore_index=True, sort=False)
        df = df[pd.notnull(df[na_col])]
        df = df.iloc[:, : defects_cols]
        # ~ convert nasty DATETIME columns into TEXT, because sqlite3 does not like them
        df[defects_str_col1] = df[defects_str_col1].astype(str)
        df[defects_str_col2] = df[defects_str_col2].astype(str)
        # ~ fill na for rest of empty values to cater sqlite3
        df.fillna(_na, inplace=True)
        if self.fix_defects_setname:
            df[self.fix_defects_setname].replace(self.fix_defects_setname_mapping, inplace=True)
            df[defects_col_FYMM] = df[self.fix_defects_setname].apply(lambda x: self._match_fymod(x, self.fix_inputs_fy_mapping))
        # ~ rename all stupid headers to fit sqlite3 schema
        df.rename(columns=new_column_names, inplace=True)
        df[defects_tv_plants].replace(defects_tv_plants_mapping, inplace=True)
        self._dfDefect = df

    def _clean_dfInputs(self, DataFrameList:List[DataFrame])->None:
        """clean inputs DataFrame before dumping into database

        :param DataFrameList: a list contains raw dataframe
        :type DataFrameList: List[DataFrame]
        """
        # ~ Inputs
        inputs_Modname:str  = 'ModelName'
        inputs_weekcode:str = 'FY_MM'

        df:DataFrame = pd.concat(DataFrameList, ignore_index=True, sort=False)
        # ~ replace stupid entries: model name, input fy, input weekcode; 20210818
        if self.fix_inputs_modelname and self.fix_inputs_modelname_mapping:
            df[self.fix_inputs_modelname].replace(self.fix_inputs_modelname_mapping, inplace=True)
        if self.fix_inputs_fy and self.fix_inputs_fy_mapping:
            df[self.fix_inputs_fy] = df[inputs_Modname].apply(lambda x: self._match_fymod(x, self.fix_inputs_fy_mapping))
        if self.fix_inputs_wkcode:
            df[inputs_weekcode] = df[self.fix_inputs_wkcode].astype(int).apply(lambda x: utils.wkno2ym(int(float(x)), datetime.datetime.today().year))
        self._dfInputs = df

    @timer
    def work(self)->None:
        sheet_name:str = 'inputs'
        df_defects = []
        df_inputs  = []

        for file in self._filter():
            tmp_dfDefects = pd.read_excel(file)
            tmp_dfInput   = pd.read_excel(file, sheet_name=sheet_name)
            df_defects.append(tmp_dfDefects)
            df_inputs.append(tmp_dfInput)
        self._clean_dfDefects(df_defects)
        self._clean_dfInputs(df_inputs)

        # ~ migrate to db
        sql_table_name_defects:str = 'defects'
        sql_table_name_inputs:str  = 'inputs'
        with sqlite3.connect(self._db) as conn:
            cur = conn.cursor()
            info = 'update {}'.format(sql_table_name_defects)
            with self.__enter_defects_table(sql_table_name_defects, cur, info):
                self._dfDefect.to_sql(sql_table_name_defects, conn, if_exists='append', index=False)
            info = 'update {}'.format(sql_table_name_inputs)
            with self.__enter_inputs_table(sql_table_name_inputs, cur, info):
                self._dfInputs.to_sql(sql_table_name_inputs, conn, if_exists='append', index=False)


if __name__ == "__main__":
    folder = r'D:\pj_00_codelib\2019_pypj\20210805 QA_weekly_report\data'
    db     = 'pmod_smld.db'
    db_web = r'D:\pj_00_codelib\2019_pypj\20210823 django-tvqa-master\db.sqlite3'

    defects_setname          = "Set name"
    defects_setname_mapping  = {'75AR(2 Stand)':'75AR', '55NH(1)':'55NH', '75NB BOE':'75NB'}
    defects_setname_mapping.update(setname_mapping.fsk_setname_mapping)

    inputs_modelname         = "ModelName"
    inputs_modelname_mapping = {'YDBM075DCS02':'75AR', '75NB BOE':'75NB', '75NXB CSOT':'75NXB', 'SB2H':'SB85', 'SBL2':'SBL49'}
    inputs_fymod         = 'FY_mod'
    inputs_fymod_mapping = {'A':'FY21', 'N':'FY20', 'S':'FY19'}
    inputs_wkcode        = "WeekCode"

    sm = Merger(folder=folder, db=db,
            fix_defects_setname=defects_setname, fix_defects_setname_mapping=defects_setname_mapping,
            fix_inputs_modelname=inputs_modelname, fix_inputs_modelname_mapping=inputs_modelname_mapping,
            fix_inputs_fy=inputs_fymod, fix_inputs_fy_mapping=inputs_fymod_mapping,
            fix_inputs_wkcode=inputs_wkcode
    )
    sm.work()
