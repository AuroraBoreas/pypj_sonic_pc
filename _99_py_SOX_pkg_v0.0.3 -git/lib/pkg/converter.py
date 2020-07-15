"""A module to convert a 32bit binary file to a 16bit binary file 

This module consists of the following basic functionalities.
- Open a binary file and return the length of it
- Read a binary file and return as an array
- Extract 16bit from 32bit array and return as a list
- Save list as a new binary file

Changelong
- v0.0.1, initial version to convert 32bit to 16bit
- v0.0.2, modify algorithm
- v0.0.3, dynamically calculate the length of a binary file

Author
@ZL, 20200628

Warning
This module is open-source and free. The result is NOT guaranteed.
Use at your own risk.

"""

import array, os

_TYP = 'B' #found during data exploration inside of source .bin files, bytes data

def get_bin_len_of_file(src_file):
    """Open a binary file and calculate the length of it.

    :param src_file: name or path of a binary file
    :type src_file: string
    """
    import numpy as np 
    with open(src_file, 'rb') as f:
        tmp_arr = np.fromfile(f)
    return len(tmp_arr)

def read_src_file(src_file):
    """Read a binary file, and return as array.array()

    :param src: file name or path of the binary file
    :type src: string
    """
    num_lon = 8 #found during data exploration inside of source .bin files, bytes data
    num_lat = get_bin_len_of_file(src_file)
    with open(src_file, 'rb') as f:
        tmp_arr = array.array(_TYP)
        tmp_arr.fromfile(f, num_lon*num_lat)
    return tmp_arr

def extract_16_from_32(arr):
    """Extract 16bit data from 32bit array

    :param arr: an array of 32bit bytes
    :type arr: list

                            0    1    2    3    4    5    6    7
    >>> extract_16_from_32(0100 1000 0001 0001 0101 1001 0101 0101)
          0     2     4     6
    >>> [0100, 0001, 0101, 0101]
    """
    start = 0
    stop = len(arr)
    step = 4    
    tmp_list = []
    for i in range(start, stop, step):
        tmp_list.extend(arr[i:i+2])
    return tmp_list

def save_bytes_to_file(seq, new_file):
    """convert list to array.array() then save as a binary file

    :param seq: a list of bytes
    :type seq: list
    """
    with open(new_file, 'wb') as f:
        tmp_arr = array.array(_TYP)
        tmp_arr.fromlist(seq)
        tmp_arr.tofile(f)
    return

def convert_32bit_to_16bit2(src_file, new_file):
    """
    Author: @SHES
    """
    index = 0
    with open(src_file, 'rb') as sf, open(new_file, 'wb') as nf:
        while True:
            buf = sf.read(2) 
            if not buf: 
                break
            if(index % 2 == 0): 
                nf.write(buf) 
            index += 1 
    os.remove(src_file)
    file_name = "aec_loopback_post.bin"
    os.rename(new_file, os.path.join(os.path.split(new_file)[0], file_name))
    return

def convert_32bit_to_16bit(src_file, new_file):
    """Convert a binary file that is 32bit to 16bit

    :param src_file: file name or path of a source binary file
    :type src_file: string
    :param new_file: file name or path of a new binary file
    :type new_file: string
    """
    # read a source file as binary
    array_32bit = read_src_file(src_file)
    # extract
    list_16bit = extract_16_from_32(array_32bit)
    # encode combined data, then write into a binary file
    save_bytes_to_file(list_16bit, new_file)
    print('Done')
    return