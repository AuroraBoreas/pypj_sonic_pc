import sys
sys.path.append('.')
from lib.core import Merger, logging
from lib.vba import caller
from lib import setname_mapping
from lib.query import Exporter

if __name__ == "__main__":
    preprocess_xl = r'D:\pj_00_codelib\2019_pypj\20210805 WW_Pmod_SMLD_Control_System\lib\vba\WWPmodSMLDWrangler.xlsm'
    macro_name = "Module1.CPModWrangler"

    folder = r'D:\pj_00_codelib\2019_pypj\20210805 WW_Pmod_SMLD_Control_System\data'
    dst_xl = 'pmod_smld.xlsx'
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

    try:
        caller.call_vba_macro(preprocess_xl, macro_name)
        logging.info("preprocess operation finished")
        try:
            sm.work()
        except:
            logging.info("PythonError: failed to merge source file")
    except:
        logging.info("VBAError: failed to clean source file")

    ex = Exporter(db, dst_xl)
    ex.work()