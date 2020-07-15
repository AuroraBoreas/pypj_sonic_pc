# Implementations

This page covers most of details of this project.

## Batch Runner

Build a common module to run batch files.

```Python
#runbat.py
import subprocess

def change_src_gain(gain_value):
    """Author: SHES

    :param gain_value: gain value
    :type gain_value: float
    """
    #src_gain変更用関数
    subprocess.call("adb root", shell=True)
    subprocess.call("adb shell setenforce 0", shell=True)
    subprocess.call("adb shell chmod 777 data", shell=True)
    ##src gainの値を変更する
    subprocess.call("adb shell setprop vendor.mtk.audio.aec.ref.gain {}".format(gain_value), shell=True)

def run_batch(batch_file):
    """wrapper to run batch file

    :param batch_file: name or path of batch file
    :type batch_file: string
    """
    proc = subprocess.Popen(batch_file, shell=True)
    proc.wait()
```

## Bit Converter

Build a module to convert bytes data from 32bit to 16bit

```Python
#converter.py
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
```

## SoX Log

Build a module to log output of `sox sample.wav -n stats`

```Python
#stats.py
import subprocess, os

def get_wav_stats(wav_file):
    """get output of sox -n stats for a single .wav file and return the output log 

    :param wav_file: name or path of .wav file
    :type wav_file: bytes
    """
    rough_wav_stat = subprocess.check_output("sox {0} -n stats".format(wav_file),
                                            stderr=subprocess.STDOUT,
                                            shell=True)
    return rough_wav_stat

def save_wav_stats_to_txt(bytes_data, new_file):
    """convert bytes data into string, and save it into text file

    :param bytes_data: bytes data from a stream
    :type bytes_data: bytes
    :param new_file: path or name of a text file
    :type new_file: string
    """
    sox_stats = bytes_data.decode("utf-8")
    with open(new_file, "a") as f:
        f.writelines(sox_stats)
    return

def get_sox_stats(wav_file_mic25, wav_file_src25, log_file):
    """wrapper to generate sox stats and save to local disk
    """
    # put .wav files into a list container
    wav_files = [
        wav_file_mic25,
        wav_file_src25
        ]    
    # remove history record, because pkg.stats modules write data in append mode
    if os.path.exists(log_file):
        os.remove(log_file)
    # write new stats into the text file
    for file in wav_files:
        tmp_wav_stats = get_wav_stats(file)
        save_wav_stats_to_txt(tmp_wav_stats, log_file)
    return
```

## Log Parser

Build a module to parse SoX log and return data.

```Python
#parser.py

def get_data_from_sox_stats_txt(src_file):
    """read sox_mic_src_stats.txt and return rows starts with ("Pk lev dB", "RMS lev dB")

    :param src_file: path or name of a sox_mic_src_stats.txt file
    :type src_file: string
    """
    PklevdB_header = "Pk lev dB"
    RMSlevdB_header = "RMS lev dB"
    with open(src_file, 'r') as f:
        tmp_data = [line for line in f.readlines() if line.startswith(PklevdB_header) or line.startswith(RMSlevdB_header)]
    return tmp_data

def parse_PklevdB_and_RMSlevdB(seq):
    """parse strings inside seq and return overall values which are float type 

    :param seq: a list of strings
    :type seq: list

    >>> parse_RMSlevdB(["Pk lev dB     -26.79    -27.14    -26.79\n", "Pk lev dB     -25.34    -25.34    -25.34\n"])
    >>> [-26.79, '-25.34']
    >>> parse_RMSlevdB(["RMS lev dB    -40.72    -42.04    -39.70\n","RMS lev dB    -35.05    -35.05    -35.05\n"])
    >>> [-40.72, -35.05]
    """
    tmp_seq = []
    index_cmn_val = 3
    #remove '' inside seq
    for item in seq:
        tmp_arr = item.strip().split(" ")
        tmp_seq.append([s for s in tmp_arr if s != ''])
    return [float(tmp_item[index_cmn_val]) for tmp_item in tmp_seq]

def get_diff(seq):
    """return difference between src_25.wav and mic_25.wav, diff = (src_25.wav - mic_25.wav)

    :param seq: a list of two floats, first is mic_25.wav("RMS lev dB", "Overall"), second is src_25.wav("RMS lev dB", "Overall")
    :type seq: list
    :return: difference
    :rtype: float
    """
    return round((seq[3] - seq[1]), 2)

def get_src_Pklev(seq):
    """return Pk lev dB of src_25.wav

    :param seq: a list of Pk, RMS values
    :type seq: list
    :return: Pk lev dB of src_25.wav
    :rtype: float
    """
    return float(seq[2])

def read_sox_stats_and_get_diff_Pklev(src_file):
    """wrapper to return diff between mic_25.wav and src_25.wav

    :param src_file: path or name of a sox_mic_src_stats.txt file
    :type src_file: string
    :return: difference(mic_25.wav - src_25.wav) and src Pk lev dB
    :rtype: tuple
    """
    seq = get_data_from_sox_stats_txt(src_file)
    lst = parse_PklevdB_and_RMSlevdB(seq)
    return get_diff(lst), get_src_Pklev(lst)
```