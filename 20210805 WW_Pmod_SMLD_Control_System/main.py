import sys
sys.path.append('.')
from lib.core import Merger, logging, os
from lib.vba import caller
from lib import setname_mapping
from lib.query import Exporter

if __name__ == "__main__":
    preprocess_xl = r'lib\vba\WWPmodSMLDWrangler.xlsm'
    macro_name    = "Module1.CPModWrangler"

    folder     = r'data'
    db         = 'pmod_smld.db'
    dst_xl     = 'pmod_smld.xlsx'
    db_web     = r'\\43.98.1.18\SSVE_Division\TV品质保证部\部共通\部门\2制品品质课\CS\FY21_Model\20210805 WW_Pmod_SMLD_System\pmod_smld.db'
    dst_xl_web = r'\\43.98.1.18\SSVE_Division\TV品质保证部\部共通\部门\2制品品质课\CS\FY21_Model\20210805 WW_Pmod_SMLD_System\pmod_smld.xlsx'

    defects_setname          = "Set name"
    defects_setname_mapping  = {'75AR(2 Stand)':'75AR', '55NH(1)':'55NH', '75NB BOE':'75NB'}
    defects_setname_mapping.update(setname_mapping.fsk_setname_mapping)
    inputs_modelname         = "ModelName"
    inputs_modelname_mapping = {'YDBM075DCS02':'75AR', '75NB BOE':'75NB', '75NXB CSOT':'75NXB', 'SB2H':'SB85', 'SBL2':'SBL49'}

    inputs_fymod         = 'FY_mod'
    inputs_fymod_mapping = {'A':'FY21', 'N':'FY20', 'S':'FY19'}
    inputs_wkcode        = "WeekCode"

    sm = Merger(folder=os.path.abspath(folder), db=os.path.abspath(db_web),
            fix_defects_setname=defects_setname, fix_defects_setname_mapping=defects_setname_mapping,
            fix_inputs_modelname=inputs_modelname, fix_inputs_modelname_mapping=inputs_modelname_mapping,
            fix_inputs_fy=inputs_fymod, fix_inputs_fy_mapping=inputs_fymod_mapping,
            fix_inputs_wkcode=inputs_wkcode
    )

    logging.info("start preprocess..")
    try:
        caller.call_vba_macro(os.path.abspath(preprocess_xl), macro_name)
        logging.info("preprocess finished")
        try:
            logging.info("start Python -> SQL..")
            sm.work()
        except:
            logging.info("PythonError: failed to merge source file")
    except:
        logging.info("VBAError: failed to clean source file")

    if os.path.exists(db_web) and os.path.exists(dst_xl_web):
        ex = Exporter(db_web, dst_xl_web)
        ex.work()
    else:
        if os.path.exists(db) and os.path.exists(dst_xl):
            ex = Exporter(db, dst_xl)
            ex.work()
        else:
            raise FileNotFoundError()