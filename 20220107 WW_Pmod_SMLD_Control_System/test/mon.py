import datetime, logging
logging.basicConfig(level=logging.INFO, format='%(message)s')

def wc2mon(week_code:int):
    wc = str(week_code)
    year = int(wc[:2])
    week = int(wc[2:])
    date = datetime.datetime(year, 1, 1+(week-1)*7)
    mon  = date.month
    logging.info('weekcode:{} -> month: {}'.format(week_code, mon))

if __name__ == '__main__':
    wkcode = 2112
    wc2mon(wkcode)