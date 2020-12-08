import os, sys
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from lib.pkg import converter, parser, runbat, stats

def meas_and_get_result():
    """step1, step2, step3 and return meas result

    :param IsInitial: write inital gain if true, defaults to True
    :type IsInitial: bool, optional
    :return: Difference(mic_25.wav - src_25.wav), src_25.wav overall Pk lev dB
    :rtype: tuple
    """
    base_dir1 = r"C:\Users\ssv\Desktop\m5\bat"
    base_dir2 = r"C:\AEC\aec_dump"
    # batch file for step1
    bat_file_whitenoise_reprod = os.path.join(base_dir1, "20200416_srcadj_get_files_for_Valhalla_tencent.bat")
    # paths of two bins files generated by sox after step1, waiting for step2 to convert
    src_file = os.path.join(base_dir2, "aec_loopback_post.bin")  # 32bit data generated by step1 measurement
    new_file = os.path.join(base_dir2, "aec_loopback_tmp.bin")   # store 16bit data, it will be renamed to aec_loopback_post.bin
    # paths of two .wav files generated by sox after step3
    wav_file_mic25 = os.path.join(base_dir1, "mic_25.wav")       # mic_25.wav in ProdDesign's JIG PC
    wav_file_src25 = os.path.join(base_dir1, "src_25.wav")       # src_25.wav in ProdDegisn's JIG PC
    log_file       = os.path.join(base_dir1, "sox_mic_src_stats")# a txt file stores sox stats
    # batch file for step3
    bat_scr_make25 = os.path.join(base_dir1, "20200416_srcadj_make25_for_Valhalla_tencent.bat")

    # step1: ssv approach differs from SHES main prg
    runbat.run_batch(batch_file=bat_file_whitenoise_reprod)
    # step2, convert 32bit to 16bit
    converter.convert_32bit_to_16bit2(src_file=src_file, new_file=new_file)
    # step3, make mic_25.wav and src_25.wav
    runbat.run_batch(batch_file=bat_scr_make25)
    # get sox stats and save it as log file
    stats.get_sox_stats(wav_file_mic25=wav_file_mic25, wav_file_src25=wav_file_src25, log_file=log_file)
    return parser.read_sox_stats_and_get_diff_Pklev(src_file=log_file)