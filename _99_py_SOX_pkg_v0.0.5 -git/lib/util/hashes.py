import hashlib, os

def hash_file(file_path):
    """display hash strings of a give file path"""
    # BUF_SIZE is totally arbitrary, change for your app!
    BUF_SIZE = 65536  # lets read stuff in 64kb chunks!
    md5 = hashlib.md5()
    sha1 = hashlib.sha1()
    with open(file_path, 'rb') as f:
        while True:
            data = f.read(BUF_SIZE)
            if not data:
                break
            md5.update(data)
            sha1.update(data)
    print("MD5: {0}".format(md5.hexdigest()))
    print("SHA1: {0}".format(sha1.hexdigest()))
    return [md5.hexdigest(), sha1.hexdigest()]