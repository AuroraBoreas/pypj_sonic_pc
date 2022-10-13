"# main" 

from lib.config import tlogger_folder
from lib.tlogger import TLogWorker
from lib.jlogger import JLogWorker
from lib.util import List
import datetime

if __name__ == '__main__':
    pmod_name:str = 'FT75'
    fn:str        = '{}_{}_{}_res.csv'.format(datetime.datetime.today().strftime('%Y%m%d%H%M%S'), pmod_name, "log")
    args:List     = (tlogger_folder, 'csv', fn)
    w:TLogWorker  = TLogWorker()
    w.sort_logs(*args)