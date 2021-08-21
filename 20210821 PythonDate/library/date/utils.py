"""
this module conerts date string to week number and backward.

it has the following functionalities.
- convert date string to week number
- convert week number to year month

changelog
- v0.01, @ZL, 20210821

"""

import datetime
import logging
from typing import NewType, Tuple
Date = NewType('Date', datetime.date)

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(message)s')

def date2wkno(date_str:str)->int:
    """return week number from a date string

    Args:
        date_str (str): date string. for example: '2021-08-21 00:00:00'

    Returns:
        int: yyww. for example: 2125
    """
    pfmt = "%Y-%m-%d %H:%M:%S" # format varies according to args
    d = datetime.datetime.strptime(date_str, pfmt).date()
    y = d.year % 100
    w = d.isocalendar()[1]
    fmt = "{}{}".format(y, w)
    return int(fmt)

def wkno2ym(week_number:int, year:int)->str:
    """return short yy mm from week number

    Args:
        week_number (int): week number. format is 2132. 21 shorts for 2021, 32 shorts for week number 32

    Returns:
        AnyStr: yy mm. for example: 21 Aug

    Ref link: https://stackoverflow.com/questions/17087314/get-date-from-week-number
    """

    i = 2 # index slices week_nummber: yyww
    d = 1 # pick monday as first day
    c = year // 100 * 100 # get first year of century

    cstr = str(week_number)
    y    = int(cstr[:i]) + c
    w    = cstr[i:]
    fmt  = "{}-{}-{}".format(y, w, d)
    d    = datetime.datetime.strptime(fmt, "%Y-%W-%w") # return yyyy-mm-dd
    return d.strftime("%y %b") # return short version yy-mm

def GetBeginEndDateFromCalendarWeek(year:int, calendar_week:int)->Tuple[Date,Date]:
    """return tuple[monday, sunday] from year, calendary week. for example: 2021, 32

    Args:
        year (int): natural year
        calendar_week (int): calendar week

    Returns:
        Tuple[Date,Date]: tuple

    Ref link: https://stackoverflow.com/questions/51194745/get-first-and-last-day-of-given-week-number-in-python
    """
    monday = datetime.datetime.strptime(f'{year}-{calendar_week}-1', "%Y-%W-%w").date()
    return monday, monday + datetime.timedelta(days=6.9)

if __name__ == '__main__':
    today = "2021-06-27 09:15:32"
    logging.info(date2wkno(today))

    week_number = 2132
    d = wkno2ym(week_number, 1985)
    logging.info("week number('{0}') -> {1}".format(week_number, d))
    logging.info("start date: {}, end date: {}".format(*GetBeginEndDateFromCalendarWeek(2021, 1)))