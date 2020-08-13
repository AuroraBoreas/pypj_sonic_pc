import sys, os, shutil, time
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from lib.pkg import converter
from lib.util import hashes

if __name__ == "__main__":
    BASE_DIR = os.path.dirname(os.path.dirname(__file__))
    # remove history file
    try:
        os.remove(os.path.join(BASE_DIR, r"data\output\aec_loopback_post.bin"))
    except FileNotFoundError:
        pass
    except FileExistsError:
        pass
    src_file = os.path.join(BASE_DIR, r"data\未转\aec_loopback_post.bin")
    new_file = os.path.join(BASE_DIR, r"data\output\tmp.bin")
    # convert
    converter.convert_32bit_to_16bit2(src_file, new_file)
    # recover src file from backup
    shutil.copy(src=os.path.join(BASE_DIR, r"data\未转\bin\aec_loopback_post.bin"), dst=os.path.join(BASE_DIR, r"data\未转"))

if __name__ == "__main__":
    # compare converted result by python vs converted result by Lua program
    BASE_DIR = os.path.dirname(os.path.dirname(__file__))
    bin_file_converted_by_LUA_prg = os.path.join(BASE_DIR, r"data\已转\aec_loopback_post.bin")
    bin_file_converted_by_Py_prg  = os.path.join(BASE_DIR, r"data\output\aec_loopback_post.bin")

    list_files = [
        bin_file_converted_by_LUA_prg,
        bin_file_converted_by_Py_prg
    ]
    
    result = [hashes.hash_file(file) for file in list_files]
    assert result[0] == result[1]