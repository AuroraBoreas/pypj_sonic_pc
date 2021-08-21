"""
unittest

"""

import sys
sys.path.append('..')
from library.date.utils import date2wkno, wkno2ym, logging

if __name__ == '__main__':
    today = '27/10/20 05:23:20'
    logging.info(date2wkno(today))

    wkcode = 2019
    d = wkno2ym(wkcode, 2018)
    logging.info("week number('{0}') -> {1}".format(wkcode, d))