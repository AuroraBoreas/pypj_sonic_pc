"""A moudle reads ```sox -n stats``` output.txt and calculates difference between mic_25.wav and src_25.wav

It has the following basic functionalities
- Reads sox output.txt and returns Pk values inside it
- Parses Pk values and returns a list of two Pk values parsed
- Get overall Pk lev dB value of src_25.wav
- Calculates differences

Changelog
- v0.0.1, initial version

Author
@ZL, 20200630

"""

def get_data_from_sox_stats_txt(src_file):
    """read sox_mic_src_stats.txt and return rows starts with ("Pk lev dB", "RMS lev dB")

    :param src_file: path or name of a sox_mic_src_stats.txt file
    :type src_file: string

    >>> get_data_from_sox_stats_txt(src_file)
    >>> [
         "Pk lev dB     -26.79    -27.14    -26.79\n", "RMS lev dB    -40.72    -42.04    -39.70\n", # sox stats of mic_25.wav
         "Pk lev dB     -25.34    -25.34    -25.34\n", "RMS lev dB    -35.05    -35.05    -35.05\n"  # sox stats of src_25.wav
        ]
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

    >>> parse_PklevdB_and_RMSlevdB([
         "Pk lev dB     -26.79    -27.14    -26.79\n", "RMS lev dB    -40.72    -42.04    -39.70\n", # sox stats of mic_25.wav
         "Pk lev dB     -25.34    -25.34    -25.34\n", "RMS lev dB    -35.05    -35.05    -35.05\n"  # sox stats of src_25.wav
        ])
    >>> [
          ["Pk", "lev", "dB", "-26.79", "-27.14", "-26.79"],
          ["RMS", "lev", "dB", "-40.72", "-42.04", "-39.70"],
          ["Pk", "lev", "dB", "-25.34", "-25.34", "-25.34"],
          ["RMS", "lev", "dB", "-35.05", "-35.05", "-35.05"],
        ]
    >>> [26.79, 40.72, 25.34, 35.05]

    """
    tmp_seq = []
    index_cmn_val = 3
    #remove '' inside seq
    for item in seq:
        tmp_arr = item.strip().split(" ")
        tmp_seq.append([s for s in tmp_arr if s != ''])
    return [abs(float(tmp_item[index_cmn_val])) for tmp_item in tmp_seq]

def get_diff(seq):
    """return difference between src_25.wav and mic_25.wav, diff = (src_25.wav - mic_25.wav)

    :param seq: a list of two floats, first is mic_25.wav("RMS lev dB", "Overall"), second is src_25.wav("RMS lev dB", "Overall")
    :type seq: list
    :return: difference
    :rtype: float
    """
    return abs(round((seq[3] - seq[1]), 2))

def get_src_Pklev(seq):
    """return Pk lev dB of src_25.wav

    :param seq: a list of Pk, RMS values
    :type seq: list
    :return: Pk lev dB of src_25.wav
    :rtype: float
    """
    return abs(float(seq[2]))

def read_sox_stats_and_get_diff_Pklev(src_file):
    """wrapper to return diff between mic_25.wav and src_25.wav

    :param src_file: path or name of a sox_mic_src_stats.txt file
    :type src_file: string
    :return: difference(mic_25.wav - src_25.wav) and src Pk lev dB
    :rtype: tuple
    """
    seq = get_data_from_sox_stats_txt(src_file)
    lst = parse_PklevdB_and_RMSlevdB(seq)
    return get_diff(lst), get_src_Pklev(lst), lst