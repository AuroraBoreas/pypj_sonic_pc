
class CIE_Converter:
    """A module to convert CIE spaces(CIE1976 <-> CIE1931)  
    Author: @ZL, 20200331

    It provides the following operations.
    - convert CIE1976 space coordiates(u', v') to CIE1931 space coorinates(x, y)
    - convert CIE1931 space coorinates(x, y) to CIE1976 space coordiates(u', v')

    Refer to the standards of CIE spaces on internet
    """
    def __init__(self, a: float, b: float, color_space: str):
        self.a = a
        self.b = b
        self.color_space = color_space

    def __repr__(self):
        if self.color_space == 'xy':
            return "xy->u\'v\', u\':{:.4f}, v\':{:.4f}".format(*self.converter())
        if self.color_space == 'dudv':
            return "u\'v\'->xy, x :{:.4f}, y: {:.4f}".format(*self.converter())
        else:
            return 'incorrect color space'

    def converter(self):
        def convert_dudv_to_xy(du: float, dv: float):
            """CIE1976 space(du,dv) to CIE1931 space(x,y) @ZL"""
            y = (3*dv)/(9*du/2 - 12*dv + 9)
            x = (du/dv)*9/4*y
            return x, y

        def convert_xy_to_dudv(x: float, y: float):
            """CIE1931 space(x,y) to CIE1976 space(du,dv) @ZL"""
            du = (4*x)/(12*y - 2*x + 3)
            dv = (9*y)/(12*y - 2*x + 3)
            return du, dv

        if self.color_space == 'xy':
            return convert_xy_to_dudv(self.a, self.b)
        if self.color_space == 'dudv':
            return convert_dudv_to_xy(self.a, self.b)
        else:
            return

if __name__ == '__main__':
    print("{:-^40}".format("CIE color coordinates convert"))
    #<~ du, dv to x, y
    du = 0.189988903970067
    dv = 0.44169542168743
    c = CIE_Converter(du, dv, color_space='dudv')
    print(c)
    #<~ x, y to du, dv
    x, y = 0.2815667, 0.2909333
    c = CIE_Converter(x, y, color_space='xy')
    print(c)
    #<~ incorrect color space
    c = CIE_Converter(x, y, color_space='abcs')
    print(c)
