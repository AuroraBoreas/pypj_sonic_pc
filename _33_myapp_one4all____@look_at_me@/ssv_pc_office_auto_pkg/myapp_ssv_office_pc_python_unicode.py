import sys, io
def change_encoding_for_py():
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

if __name__ == '__main__':
    change_encoding_for_py()
