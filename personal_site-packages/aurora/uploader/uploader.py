import socket, sys, os
from distutils import dir_util

class Intranet_Uploader:
    def __init__(self, src_dir, dst_dir, host = '43.98.1.18', port = 445):
        self.src_dir = src_dir
        self.dst_dir = os.path.join(dst_dir, os.path.split(src_dir)[-1])
        self.host    = host
        self.port    = port

    def has_intranet(self, host, port, timeout=3):
        """check intranect connectivity"""
        try:
            socket.setdefaulttimeout(timeout)
            socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect((host, port))
            return True
        except socket.error:
            return False

    def start(self):
        if self.has_intranet(self.host, self.port):
            dir_util.copy_tree(self.src_dir, self.dst_dir)
            return True
        else:
            sys.exit("Warning: Lost SSV intranet connection")
            return False

if __name__ == "__main__":
    sd = r"C:\Users\5106001995\Desktop\20210121 dummy"
    dd = r"\\43.98.1.18\SSVE_Division\SHES-C\部共通\03-Info-Share\03-FYxx Model\FY21_Model\AG\21.CS"
    iu = Intranet_Uploader(sd, dd)
    iu.start()