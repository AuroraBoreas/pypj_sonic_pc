import datetime

if __name__ == '__main__':
    fn:str = str(datetime.datetime.now()).replace(' ', '_')[:19] + 'res'
    print(fn)