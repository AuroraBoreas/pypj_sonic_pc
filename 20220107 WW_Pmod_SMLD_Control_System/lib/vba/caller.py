"""
this module enables running excel macro via python.
it bridges communication and integration python and vba.

functionalities
- get excel application instance
- get workbook instance
- run marco in the workbook instance

author
- @ZL, 20210821

changelog
- v0.01, initial build

"""

import os, time, sys
sys.path.append('.')
import contextlib
import traceback
import win32com.client
import pywintypes
import functools
from typing import NewType, Callable, Any
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(message)s')
Path = NewType('Path', str)

def timer(func:Callable)->Callable:
    @functools.wraps(func)
    def inner(*args:Any, **kwargs:Any)->Any:
        beg = time.perf_counter()
        rv  = func(*args, **kwargs)
        end = time.perf_counter()
        logging.info('Successed. time lapsed(s): {0:.2f}'.format(end - beg))
        return rv
    return inner

@contextlib.contextmanager
def OpenExcel():
    """return excel application object

    Yields:
        [type]: excel application object

    Ref link: https://stackoverflow.com/questions/19616205/running-an-excel-macro-via-python
    """
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        yield excel
    finally:
        # excel.Application.Save() # if you want to save then uncomment this line and change delete the ", ReadOnly=1" part from the open function.
        excel.Application.Quit()
        del excel

@contextlib.contextmanager
def OpenWorkbook(excel, xl_file):
    """return workbook object via excel application

    Args:
        excel ([type]): excel application object
        xl_file ([type]): excel file(*.xlsm) that contains macro

    Yields:
        [type]: workbook object

    Ref link: https://stackoverflow.com/questions/49904045/win32com-save-as-variable-name-in-python-3-6
    """
    xl_file  = os.path.expanduser(xl_file)
    workbook = excel.Workbooks.Open(Filename=xl_file, ReadOnly=1)
    try:
        yield workbook
    finally:
        # workbook.Save()
        workbook.Close(False)

@timer
def call_vba_macro(xl_file:Path, macro_name:str)->None:
    """call vba macro to complement miss data and summarize inputs, due to time efficiency
    vba is much faster than python + xlrd
    """
    # xl_file    = r"D:\pj_00_codelib\2019_pypj\20210805 QA_weekly_report\lib\vba\WWPmodSMLDWrangler.xlsm"
    # macro_name = "Module1.CPModWrangler"
    macro_addr = "\'{0}\'!{1}".format(xl_file, macro_name)
    if os.path.exists(xl_file):
        with OpenExcel() as excel:
            try:
                with OpenWorkbook(excel, xl_file) as workbook:
                    # possible workbook usage
                    excel.Application.Run(macro_addr) # run macro in the workbook
            except pywintypes.com_error as e:
                # print(xl_file)
                traceback.print_exc()

if __name__ == '__main__':
    # ~ xl_file must be an absolute path @ZL, 20210821
    preprocess_xl = r'D:\pj_00_codelib\2019_pypj\20210805 WW_Pmod_SMLD_Control_System\lib\vba\WWPmodSMLDWrangler.xlsm'
    macro_name = "Module1.CPModWrangler"
    call_vba_macro(preprocess_xl, macro_name)