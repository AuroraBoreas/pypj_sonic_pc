
class CIE_Converter:
    """A module to convert CIE spaces(CIE1976 <-> CIE1931)  
    Author: @ZL, 20200331

    It provides the following operations.
    - convert CIE1976 space coordiates(u', v') to CIE1931 space coorinates(x, y)
    - convert CIE1931 space coorinates(x, y) to CIE1976 space coordiates(u', v')

    Refer to the standards of CIE spaces on internet

    # changelog
    - v.01, 20200331, initial build
    - v.02, 20210612, refactor

    """
    _cs1931 = 'cie1931'
    _cs1976 = 'cie1976'
    _cs1931_alias = 'xy'
    _cs1976_alias = 'u\'v\''
    _precision = 4

    def __init__(self, a: float, b: float, color_space: str):
        """init

        :param a: x or u'
        :type a: float
        :param b: y or v'
        :type b: float
        :param color_space: CIE1931(x, y) or CIE1976(u', v')
        :type color_space: str
        """
        self.a = a
        self.b = b
        self.cs = color_space.lower()

    def __repr__(self):
        return "{:>7}: ({:.4f}, {:.4f})".format(self.cs, self.a, self.b)

    def __dudv_to_xy(self, du: float, dv: float)->tuple:
        """CIE1976 space(du,dv) to CIE1931 space(x,y) @ZL"""
        y = (3*dv)/(9*du/2 - 12*dv + 9)
        x = (du/dv)*9/4*y
        return format(self._cs1931_alias, ">4"), round(x, self._precision), round(y, self._precision)

    def __xy_to_dudv(self, x: float, y: float)->tuple:
        """CIE1931 space(x,y) to CIE1976 space(du,dv) @ZL"""
        du = (4*x)/(12*y - 2*x + 3)
        dv = (9*y)/(12*y - 2*x + 3)
        return format(self._cs1976_alias, ">4"), round(du, self._precision), round(dv, self._precision)
        
    def convert(self)->tuple:
        if self.cs in (self._cs1931, self._cs1931_alias):
            return self.__xy_to_dudv(self.a, self.b)
        if self.cs in (self._cs1976, self._cs1976_alias):
            return self.__dudv_to_xy(self.a, self.b)
        else:
            raise ValueError("incorrect color space")

if __name__ == '__main__':
    #<~ du, dv to x, y
    du = 0.189988903970067
    dv = 0.44169542168743
    c1 = CIE_Converter(du, dv, color_space='CIE1976')
    print(c1, c1.convert())
    #<~ x, y to du, dv
    x, y = 0.2815667, 0.2909333
    c2 = CIE_Converter(x, y, color_space='Xy')
    print(c2, c2.convert())
    #<~ incorrect color space
    # c = CIE_Converter(x, y, color_space='abcs')
    # print(c)
