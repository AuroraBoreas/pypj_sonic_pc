from typing import NewType, Dict, Any, Union, List, Callable, Tuple
import logging
from pandas import DataFrame
import pandas as pd
from sqlite3 import Cursor
import sqlite3
from contextlib import contextmanager
import functools
import time
import datetime
import os

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(message)s')

Date = NewType('Date', datetime.date)
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
