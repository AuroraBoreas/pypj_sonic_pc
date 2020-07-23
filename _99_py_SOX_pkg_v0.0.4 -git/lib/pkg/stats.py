"""A module to get output of command 'sox -n stats'

It has the following functionalities.
- Get stats of a .wav file
- Save the result into a .txt file
- Wrapper to get sox stats with one shot

Changelong
- v0.0.1, initial version to get stats of .wav file
- v0.0.2, get stats of the two .wav files.

Author
@ZL, 20200629

"""

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