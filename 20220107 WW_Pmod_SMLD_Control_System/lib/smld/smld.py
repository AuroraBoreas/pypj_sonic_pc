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

from __future__ import annotations
from abc import ABC, abstractmethod
from lib.utility.utils import wkno2ym
from lib.utility.types import (
    Dict, Any, Union, List,
    Path,
    DataFrame, Cursor,
    contextmanager,
    logging,
    pd,
    datetime,
    sqlite3,
)
import os

class IBuilder(ABC):
    @property
    @abstractmethod
    def smld(self)-> Smld:
        pass

    @abstractmethod
    def work(self)->None:
        pass

class Builder(IBuilder):
    def __init__(self, folder:Path, db:Path) -> None:
        self._folder = folder
        self._db     = db

    @property
    def smld(self) -> Smld:
        return self._smld

    @smld.setter
    def smld(self, val:Smld)->None:
        self._smld = val

    def work(self) -> None:
        self._smld.clean(self._folder)
        self._smld.to_sql(self._db)
        self._smld.reset()

class Smld:
    def __init__(self,
                fix_defects_setname:str=None, fix_defects_setname_mapping:Dict[Any, Any]=None,
                fix_inputs_modelname:str=None, fix_inputs_modelname_mapping:Dict[Any,Any]=None,
                fix_inputs_fy:str=None, fix_inputs_fy_mapping:Dict[Any,Any]=None
        ):
        """read preprocessed source excel files(*.xlsx) from a given directory, and dump into database

        Args:
            fix_defects_setname (str, optional): [description]. Defaults to None.
            fix_defects_setname_mapping (Dict[Any, Any], optional): [description]. Defaults to None.
            fix_inputs_modelname (str, optional): [description]. Defaults to None.
            fix_inputs_modelname_mapping (Dict[Any,Any], optional): [description]. Defaults to None.
            fix_inputs_fy (str, optional): [description]. Defaults to None.
            fix_inputs_fy_mapping (Dict[Any,Any], optional): [description]. Defaults to None.
        """

        self.fix_defects_setname          = fix_defects_setname
        self.fix_defects_setname_mapping  = fix_defects_setname_mapping
        self.fix_inputs_modelname         = fix_inputs_modelname
        self.fix_inputs_modelname_mapping = fix_inputs_modelname_mapping
        self.fix_inputs_fy                = fix_inputs_fy
        self.fix_inputs_fy_mapping        = fix_inputs_fy_mapping
        self.reset()
    
    def _filter(self, src_folder:Path):
        """filter source files, especially excel temporary files which start with '~$'

        :yield: excel file
        :rtype: path-like string
        """
        xl_ext:str = '.xlsx'
        tmp_xl:str = '~$'
        idns:list  = ['smld', 'quality']

        for root, _, files in os.walk(src_folder):
            for file in files:
                ext = os.path.splitext(file)[-1]
                if any(idn in file.lower() for idn in idns) and not file.startswith(tmp_xl) and xl_ext == ext:
                    logging.info(file)
                    yield os.path.abspath(os.path.join(root, file))

    @contextmanager
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

                UNIQUE(ConfirmationWC, TVPlant, SetName, ModelName, ModuleID) ON CONFLICT IGNORE
            );
            """.format(name))
        try:
            yield
        finally:
            logging.info(f'SQL {act} completed')

    @contextmanager
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
        defects_cols:int = 21
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
        # df = df[pd.notnull(df[na_col])]
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
        inputs_Modname:str       = 'ModelName'
        inputs_FYMM:str          = 'FY_MM'
        inputs_header_wkcode:str = "WeekCode"

        df:DataFrame = pd.concat(DataFrameList, ignore_index=True, sort=False)
        # ~ replace stupid entries: model name, input fy, input weekcode; 20210818
        if self.fix_inputs_modelname and self.fix_inputs_modelname_mapping:
            df[self.fix_inputs_modelname].replace(self.fix_inputs_modelname_mapping, inplace=True)
        if self.fix_inputs_fy and self.fix_inputs_fy_mapping:
            df[self.fix_inputs_fy] = df[inputs_Modname].apply(lambda x: self._match_fymod(x, self.fix_inputs_fy_mapping))
        
        df[inputs_FYMM] = df[inputs_header_wkcode].astype(int).apply(lambda x: wkno2ym(int(float(x)), datetime.datetime.today().year)) # reverse workcode to FYMM
        self._dfInputs = df

    def clean(self, folder:Path)->None:
        sheet_name:str = 'inputs'

        for file in self._filter(folder):
            tmp_dfDefects = pd.read_excel(file)
            tmp_dfInput   = pd.read_excel(file, sheet_name=sheet_name)
            self.df_defects.append(tmp_dfDefects)
            self.df_inputs.append(tmp_dfInput)
        self._clean_dfDefects(self.df_defects)
        self._clean_dfInputs(self.df_inputs)

    def to_sql(self, db:Path)->None:
        # ~ migrate to db
        sql_table_name_defects:str = 'defects'
        sql_table_name_inputs:str  = 'inputs'
        with sqlite3.connect(db) as conn:
            cur = conn.cursor()
            info = 'update {}'.format(sql_table_name_defects)
            with self.__enter_defects_table(sql_table_name_defects, cur, info):
                self._dfDefect.to_sql(sql_table_name_defects, conn, if_exists='append', index=False)
            info = 'update {}'.format(sql_table_name_inputs)
            with self.__enter_inputs_table(sql_table_name_inputs, cur, info):
                self._dfInputs.to_sql(sql_table_name_inputs, conn, if_exists='append', index=False)

    def reset(self)->None:
        self._dfDefect:DataFrame = None
        self._dfInputs:DataFrame = None
        self.df_defects:List[DataFrame] = []
        self.df_inputs:List[DataFrame]  = []