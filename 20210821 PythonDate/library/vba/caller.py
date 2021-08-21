"""
this module enables running excel macro via python.
it bridges communication and integration python and vba.

functionalities
- get excel application instance
- get workbook instance
- run marco in the workbook instance

changelog
- v0.01, initial build; @ZL 20210821

"""

import os
import contextlib
import traceback
import win32com.client
# from pythoncom import com_error
import pywintypes

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

if __name__ == '__main__':
    xl_file    = r"C:\Users\Aurora_Boreas\Desktop\pypj_sonic_pc\20210821 PythonDate\data\test1.xlsm"
    xl_delimit = "!"
    macr_name  = "Module1.Macro1"
    macro_addr = "{0}{1}{2}".format(xl_file, xl_delimit, macr_name)

    if os.path.exists(xl_file):
        with OpenExcel() as excel:
                try:
                    with OpenWorkbook(excel, xl_file) as workbook:
                        excel.Application.Run(macro_addr)
                except pywintypes.com_error as e:
                    print(xl_file)
                    traceback.print_exc()