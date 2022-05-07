import pathlib

def main():
    for i in range(1, 31):
        print("12494000-1500{:0>3}".format(i))

def parent_dir():
    print(pathlib.Path(__name__).parent.parent.absolute())

if __name__ == "__main__":
    parent_dir()