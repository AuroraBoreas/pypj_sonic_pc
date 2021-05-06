"Python is a protocol orientated lang; every top-level function has a corresponding dunder method implemented;" 

class Base:
    def foo(self):
        return self.bar()

    def __init_subclass__(cls, *args, **kwargs):
        print('__init_subclass__ : ', cls, args, kwargs)
        assert hasattr(cls, 'bar'), "bar() not found in user class!"
        return super().__init_subclass__()