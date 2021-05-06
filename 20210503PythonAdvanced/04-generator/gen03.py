"#Python is a protocol orientated lang; every top-level function has a corresponding dunder method implemented;" 

import time

def compute(n: int)->int:
    for i in range(n):
        yield i
        time.sleep(.5)

class API:
    def func1(self):
        print("func1 does stuff")
    def func2(self):
        print("func2 does stuff")
    def func3(self):
        print("func3 does stuff")
    def func4(self):
        print("func4 does stuff")

    def __call__(self):
        self.func1()
        yield
        self.func2()
        yield
        self.func3()
        yield
        self.func4()
        yield


if __name__ == '__main__':
    for i in compute(10):
        print(i)


    print('\n')
    api: API = API()
    for f in api():
        print('control in main')