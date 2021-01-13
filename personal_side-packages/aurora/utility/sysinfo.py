import platform, socket, re, uuid, json, psutil, logging, sys, ctypes

def getSystemInfo():
    """a function retrieves system info

    :return: None
    :rtype: None
    """
    try:
        info = {}
        info['platform']         = platform.system()
        info['platform-release'] = platform.release()
        info['platform-version'] = platform.version()
        info['architecture']     = platform.machine()
        info['hostname']         = socket.gethostname()
        info['ip-address']       = socket.gethostbyname(socket.gethostname())
        info['mac-address']      = ':'.join(re.findall('..', '%012x' % uuid.getnode()))
        info['processor']        = platform.processor()
        info['ram']              = str(round(psutil.virtual_memory().total / (1024.0 **3)))+" GB"
        info['python -V']        = sys.version
        info['screensize']       = ctypes.windll.user32.GetSystemMetrics(0), ctypes.windll.user32.GetSystemMetrics(1)
        return json.dumps(info)
    except Exception as e:
        logging.exception(e)

if __name__ == "__main__":
    info = json.loads(getSystemInfo())

    for k,v in info.items():
        # Markdown output
        print("* {0:<16} : {1:}".format(k, v))