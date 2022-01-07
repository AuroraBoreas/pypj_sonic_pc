# import sys
# sys.path.append('.')
from lib.smld.smld import Builder, Smld
from lib.utility.types import timer

class Director:
    def __init__(self):
        self._builder = None

    @property
    def builder(self)->Builder:
        return self._builder

    @builder.setter
    def builder(self, val:Builder)->None:
        self._builder = val

    @timer
    def start(self)->None:
        self._builder.work()


if __name__ == "__main__":
    from lib.config.config import (
        folder,
        db_web,
    )

    from lib.config.setname_mapping import (
        common_defects_setname,
        common_defects_setname_mapping,
        common_inputs_fymod,
        common_inputs_fymod_mapping,
        common_inputs_modelname,
        common_inputs_modelname_mapping
    )

    sm = Smld(fix_defects_setname=common_defects_setname, fix_defects_setname_mapping=common_defects_setname_mapping,
            fix_inputs_modelname=common_inputs_modelname, fix_inputs_modelname_mapping=common_inputs_modelname_mapping,
            fix_inputs_fy=common_inputs_fymod, fix_inputs_fy_mapping=common_inputs_fymod_mapping,
    )

    b = Builder(folder, db_web)
    d = Director()
    d.builder = b
    d.work()