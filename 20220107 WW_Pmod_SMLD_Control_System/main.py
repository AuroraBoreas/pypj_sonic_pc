import sys
sys.path.append('.')

from lib.core import Director, Smld, Builder
from lib.utility.types import logging, os
from lib.vba import caller
from lib.query import query

from lib.config.config import (
    preprocess_xl,
    macro_name,
    dst_xl,
    folder,
    db,
    db_web,
    dst_xl_web,
)

from lib.config.setname_mapping import (
    common_defects_setname,
    common_defects_setname_mapping,
    common_inputs_fymod,
    common_inputs_fymod_mapping,
    common_inputs_modelname,
    common_inputs_modelname_mapping
)

def main()->None:
    smld = Smld(
        fix_defects_setname=common_defects_setname, fix_defects_setname_mapping=common_defects_setname_mapping,
        fix_inputs_modelname=common_inputs_modelname, fix_inputs_modelname_mapping=common_inputs_modelname_mapping,
        fix_inputs_fy=common_inputs_fymod, fix_inputs_fy_mapping=common_inputs_fymod_mapping,
    )

    b = Builder(folder, db_web)
    b.smld = smld
    d = Director()
    d.builder = b

    logging.info("start preprocess..")
    try:
        caller.call_vba_macro(os.path.abspath(preprocess_xl), macro_name)
        logging.info("preprocess finished")
        try:
            logging.info("start Python -> SQL..")
            d.start()
        except:
            logging.info("PythonError: failed to merge source file")
    except:
        logging.info("VBAError: failed to clean source file")

    if os.path.exists(db_web):
        ex = query.Exporter(db_web, dst_xl_web)
        ex.work()
    else:
        if os.path.exists(db):
            ex = query.Exporter(db, dst_xl)
            ex.work()
        else:
            raise FileNotFoundError()

if __name__ == "__main__":
    main()