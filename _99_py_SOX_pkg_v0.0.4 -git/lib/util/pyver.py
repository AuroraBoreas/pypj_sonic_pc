import sys

def is_64bit() -> bool:
    return sys.maxsize > 2**32