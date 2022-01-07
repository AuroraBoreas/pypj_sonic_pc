"""
this module handles queries against pmod_smld.db;

it has the following functionalities
- add more unique-combination-columns for "defects" and "inputs" tables
- save the queries into excel for further visualization usage

author
- @ZL, 20210827

changelog
- v0.01, initial build

"""

from lib.utility.types import (
    Path, DataFrame, 
    datetime, timer, 
    pd, sqlite3,
)

from lib.utility import utils

class Exporter:
    _tbl_name_defect:str = 'defects'
    _tbl_name_inputs:str = 'inputs'

    def __init__(self, db:Path, dst_xl:Path):
        self._db       = db
        self._dst_xl   = dst_xl
        self._dfDefect:DataFrame = None
        self._dfInputs:DataFrame = None

    def __repr__(self):
        return 'Exporter {0} -> {1}'.format(self._db, self._dst_xl)
    
    def _read_sql_tables(self)->None:
        query_defect = 'SELECT * FROM {};'.format(self._tbl_name_defect)
        query_inputs = 'SELECT * FROM {};'.format(self._tbl_name_inputs)
        with sqlite3.connect(self._db) as conn:
            self._dfDefect = pd.read_sql_query(query_defect, conn)
            self._dfInputs = pd.read_sql_query(query_inputs, conn)

    def _reject_from_line(self, x:str)->str:
        idns = ['sbi', 'iqc']
        rv1 = 'line'
        rv2 = 'na'
        return rv2 if str(x).lower() in idns else rv1
    
    def _cvt_wkcode2fymm(self, x:int)->str:
        na = 'na'
        if str(x) != na:
            return utils.wkno2ym(int(float(x)), datetime.datetime.today().year)
        else:
            return na

    def _fix_defects_confirmwc(self, x:int)->str:
        na = 'na'
        if x != na:
            return str(int(float(x)))
        else:
            return na

    def _combine_columns_defects(self)->None:
        ConfirmationWC_Classify_FY_Mod = 'ConfirmationWC_Classify_FY_Mod'
        pl_from_col1 = 'ConfirmationWC'
        pl_from_col2 = 'Classify'
        pl_from_col3 = 'FY_mod'
        pl_from_col4 = 'LineName'

        Classify_ProductLine_FYMM = 'Classify_ProductLine_FYMM'
        cpf_from_col1 = 'Classify'
        cpf_from_col2 = 'FY_mod'
        cpf_from_col3 = 'LineName'
        cpf_from_col4 = 'ConfirmationWC'

        ConfirmationWC_Classify_FY_Mod_TVPlant         = 'ConfirmationWC_Classify_FY_Mod_TVPlant'
        Classify_ProductLine_FYMM_TVPlant              = 'Classify_ProductLine_FYMM_TVPlant'
        ConfirmationWC_Classify_FY_Mod_TVPlant_Setname = 'ConfirmationWC_Classify_FY_Mod_TVPlant_Setname'
        Classify_ProductLine_FYMM_TVPlant_Setname      = 'Classify_ProductLine_FYMM_TVPlant_Setname'
        ConfirmationWC_Classify_FY_Mod_Setname         = 'ConfirmationWC_Classify_FY_Mod_Setname'
        Classify_ProductLine_FYMM_Setname              = 'Classify_ProductLine_FYMM_Setname'

        TVPlant = 'TVPlant'
        SetName = 'SetName'

        self._dfDefect[ConfirmationWC_Classify_FY_Mod] = (
            self._dfDefect[pl_from_col1].apply(self._fix_defects_confirmwc) +
            self._dfDefect[pl_from_col2] +
            self._dfDefect[pl_from_col3] +
            self._dfDefect[pl_from_col4].apply(self._reject_from_line)
        )

        self._dfDefect[Classify_ProductLine_FYMM] = (
            self._dfDefect[cpf_from_col1] +
            self._dfDefect[cpf_from_col2] +
            self._dfDefect[cpf_from_col3].apply(self._reject_from_line) +
            self._dfDefect[cpf_from_col4].apply(self._cvt_wkcode2fymm)
        )

        self._dfDefect[ConfirmationWC_Classify_FY_Mod_TVPlant] = (
            self._dfDefect[ConfirmationWC_Classify_FY_Mod] +
            self._dfDefect[TVPlant]
        )

        self._dfDefect[Classify_ProductLine_FYMM_TVPlant] = (
            self._dfDefect[Classify_ProductLine_FYMM] +
            self._dfDefect[TVPlant]
        )

        self._dfDefect[ConfirmationWC_Classify_FY_Mod_TVPlant_Setname] = (
            self._dfDefect[ConfirmationWC_Classify_FY_Mod_TVPlant] +
            self._dfDefect[SetName]
        )

        self._dfDefect[Classify_ProductLine_FYMM_TVPlant_Setname] = (
            self._dfDefect[Classify_ProductLine_FYMM_TVPlant] +
            self._dfDefect[SetName]
        )

        self._dfDefect[ConfirmationWC_Classify_FY_Mod_Setname] = (
            self._dfDefect[ConfirmationWC_Classify_FY_Mod] +
            self._dfDefect[SetName]
        )

        self._dfDefect[Classify_ProductLine_FYMM_Setname] = (
            self._dfDefect[Classify_ProductLine_FYMM] +
            self._dfDefect[SetName]
        )

    def _combine_columns_inputs(self)->None:
        WeekCode_FY_mod = 'WeekCode_FY_mod'
        wfm_from_col1   = 'WeekCode'
        wfm_from_col2   = 'FY_mod'

        Fy_mod_FY_MM  = 'FY_mod_FY_MM'
        fmf_from_col1 = 'FY_mod'
        fmf_from_col2 = 'FY_MM'

        WeekCode_FY_mod_TVPlant           = 'WeekCode_FY_mod_TVPlant'
        Fy_mod_FY_MM_TVPlant              = 'FY_mod_FY_MM_TVPlant'
        WeekCode_FY_mod_TVPlant_ModelName = 'WeekCode_FY_mod_TVPlant_ModelName'
        Fy_mod_FY_MM_TVPlant_ModelName    = 'FY_mod_FY_MM_TVPlant_ModelName'
        WeekCode_FY_mod_ModelName         = 'WeekCode_FY_mod_ModelName'
        Fy_mod_FY_MM_ModelName            = 'FY_mod_FY_MM_ModelName'

        TVPlant   = 'TVPlant'
        ModelName = 'ModelName'


        self._dfInputs[WeekCode_FY_mod] = (
            self._dfInputs[wfm_from_col1].astype(str) +
            self._dfInputs[wfm_from_col2].astype(str)
        )

        self._dfInputs[Fy_mod_FY_MM] = (
            self._dfInputs[fmf_from_col1] +
            self._dfInputs[fmf_from_col2]
        )

        self._dfInputs[WeekCode_FY_mod_TVPlant] = (
            self._dfInputs[WeekCode_FY_mod] +
            self._dfInputs[TVPlant]
        )

        self._dfInputs[Fy_mod_FY_MM_TVPlant] = (
            self._dfInputs[Fy_mod_FY_MM] +
            self._dfInputs[TVPlant]
        )

        self._dfInputs[WeekCode_FY_mod_TVPlant_ModelName] = (
            self._dfInputs[WeekCode_FY_mod_TVPlant] +
            self._dfInputs[ModelName]
        )

        self._dfInputs[Fy_mod_FY_MM_TVPlant_ModelName] = (
            self._dfInputs[Fy_mod_FY_MM_TVPlant] +
            self._dfInputs[ModelName]
        )

        self._dfInputs[WeekCode_FY_mod_ModelName] = (
            self._dfInputs[WeekCode_FY_mod] +
            self._dfInputs[ModelName]
        )

        self._dfInputs[Fy_mod_FY_MM_ModelName] = (
            self._dfInputs[Fy_mod_FY_MM] +
            self._dfInputs[ModelName]
        )

    def _export_toxl(self)->None:
        with pd.ExcelWriter(self._dst_xl) as writer:
            self._dfDefect.to_excel(writer, sheet_name=self._tbl_name_defect, index=False)
            self._dfInputs.to_excel(writer, sheet_name=self._tbl_name_inputs, index=False)

    @timer
    def work(self):
        self._read_sql_tables()
        self._combine_columns_defects()
        self._combine_columns_inputs()
        self._export_toxl()

if __name__ == '__main__':
    dst_xl = 'pmod_smld.xlsx'
    db     = 'pmod_smld.db'
    # db_web = r'D:\pj_00_codelib\2019_pypj\20210823 django-tvqa-master\db.sqlite3'
    ex = Exporter(db, dst_xl)
    ex.work()