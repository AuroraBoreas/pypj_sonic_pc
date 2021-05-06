"#Python is a protocol orientated lang; every top-level function has a corresponding dunder method implemented;" 

import time

class Compute:
    def __init__(self, last: int)->None:
        self.first = 0
        self.last  = last

    def __iter__(self):
        return self

    def __next__(self):
        rv = self.first
        self.first += 1
        time.sleep(.5)
        if self.first > self.last:
            raise StopIteration()
        return rv
        
if __name__ == '__main__':
    c1: Compute = Compute(10)
    for i in c1:
        print(i)