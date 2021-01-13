import datetime, pathlib

def is_expired(expired_date: str='Dec 1 2020 8:00AM'):
    """a function decides where current date time ixpired

    :param expired_date: a datetime expired, defaults to 'Dec 1 2020 8:00AM'
    :type expired_date: str, optional
    :return: return true if expired or false if not
    :rtype: bool
    """
    ed = datetime.datetime.strptime(expired_date,'%b %d %Y %I:%M%p')
    now = datetime.datetime.now()
    return (now - ed) > datetime.timedelta(days=1)

def is_outdated(file_path: str, days: int=1):
    mtime = pathlib.Path(file_path).stat().st_mtime
    mtime = datetime.datetime.fromtimestamp(mtime)
    now = datetime.datetime.now()
    return (now - mtime) > datetime.timedelta(days=days)