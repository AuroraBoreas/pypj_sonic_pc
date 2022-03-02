"""
this module has a simple functionality to create setname mapping

changelog
- v0.01
- v0.02

author
- @ZL, 20210824

"""

ctti_setname_mapping = {
    '75AR(2 Stand)' : '75AR'
}

fsk_setname_mapping = {
    'AG65_YDBO065DBU02'    : '65AG',
    'AG75_YDBO075DAS02'    : '75AG',
    'AG85_YDBO085DAU02'    : '85AG',
    'APH75_YSBM075CNO02'   : '75APH',
    'AQ75_HDBO075DDU02'    : '75AQ',
    'AQ85_HDBO085DDU02'    : '85AQ',
    'AR75_YDBM075DCS02_2S' : '75AR',
    'AR75_YDBM075DCS02_4S' : '75AR',
    'AX75_YDBM075CCS02_2S' : '75AX',
    'AX75_YDBM075CCS02_4S' : '75AX',
    'AX85_YDBM085CNU02'    : '85AX',
    'NB75_YSAS075CNG02_S'  : '75NB',
    'NB75_YSAS075CNO02_B'  : '75NB',
    'NB85_YSAS085CNU02'    : '85NB',
    'NH55_YDAS055DNU02'    : '55NH',
    'NH65_YDAS065DNU02'    : '65NH',
    'NH75_YDAS075DNN02'    : '75NH',
    'NX75_YDAS075DND02_S'  : '75NX',
    'NX75_YDAS075DNS02_C'  : '75NX',
    'NX85_YDAS085DNU01'    : '85NX',
    'NXB75_YDAS075DNS02_C' : '75NXB',
    'NXB85_YDAS085DNU01'   : '85NXB',
    'SB2H_YD9S085DTU01'    : '85SB2H',
    'SBL2_YD9S049DND01'    : '49SBL2',
}

common_defects_setname         = "Set name"
common_defects_setname_mapping = {
    '55NH(1)'  : '55NH',
    '75NB BOE' : '75NB',
}
common_defects_setname_mapping.update(fsk_setname_mapping)
common_defects_setname_mapping.update(ctti_setname_mapping)

common_inputs_modelname         = "ModelName"
common_inputs_modelname_mapping = {
    # FY21
    'YDBM075DCS02' : '75AR',
    '75NB BOE'     : '75NB',
    '75NXB CSOT'   : '75NXB',
    'SB2H'         : 'SB85',
    'SBL2'         : 'SBL49',
    # FY22
    'HDCO075MDS02' : '75FW',
    'HDCO085MDU02' : '85FW',
    'YDCS065MES02' : '65FM',
    'YDCS075MES02' : '75FM',
    'YDCS085MDU02' : '85FM',
    'YDCV055DCS02' : '55FT',
    'YDCV065DCS02' : '65FT',
    'YDCM075DCS02' : '75FT',
    'YDCV075DCS02' : '75FT',
    'YDCM085DCU02' : '85FT',
    'YDCM065CCS02' : '65FH',
    'YDCM075CCS02' : '75FH',
    'YDCM085CCU02' : '85FH',
    'YSCM075CCO02' : '75FE',

}

common_inputs_fymod         = 'FY_mod'
common_inputs_fymod_mapping = {
    'A':'FY21', 
    'N':'FY20', 
    'S':'FY19',
}

def fsk_setname_map():
    txt = """
AG65_YDBO065DBU02
AG75_YDBO075DAS02
AG85_YDBO085DAU02
APH75_YSBM075CNO02
AQ75_HDBO075DDU02
AQ85_HDBO085DDU02
AR75_YDBM075DCS02_2S
AR75_YDBM075DCS02_4S
AX75_YDBM075CCS02_2S
AX75_YDBM075CCS02_4S
AX85_YDBM085CNU02
NB75_YSAS075CNG02_S
NB75_YSAS075CNO02_B
NB85_YSAS085CNU02
NH55_YDAS055DNU02
NH65_YDAS065DNU02
NH75_YDAS075DNN02
NX75_YDAS075DND02_S
NX75_YDAS075DNS02_C
NX85_YDAS085DNU01
NXB75_YDAS075DNS02_C
NXB85_YDAS085DNU01
SB2H_YD9S085DTU01
SBL2_YD9S049DND01
    """
    import re
    # pattern = re.compile(r'\d\d')
    for i in txt.strip().splitlines():
        rv = i.split('_')[0]
        match  = re.search(r'\d\d', rv)
        if match:
            size = match.group(0)
            pmod = rv.split(size)[0]
            fmt = "\'{0}\' : \'{1:>}\',".format(i, size+pmod)
            print(fmt)
        else:
            fmt = "\'{0}\' : \'{1:>}\',".format(i, rv)
            print(fmt)



if __name__ == '__main__':
    fsk_setname_map()