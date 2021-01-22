import os, glob, datetime, collections
import pandas as pd

def merge_LCM_log(fd_path: str):
    """a function merges all LCM log files into one csv file.

    :param fd_path: a folder that stores source csv files
    :type fd_path: the folder path
    """
    if os.path.isdir(fd_path):
        files = glob.glob(os.path.join(fd_path, "*.csv"))
        files.sort()
        n = 84
        col_names = ['h' + str(i) for i in range(n)]
        big_df = pd.concat((pd.read_csv(f, header=None, usecols=range(n), names=col_names, engine='python') \
                                        for f in files), ignore_index=True, sort=False)

        big_df = big_df.dropna(subset=['h' + str(n - 1)])
        fn = os.path.join(fd_path, 'all_LCM_merged.xlsx')
        big_df.to_excel(fn)

def merge_SET_log(fd_path: str):
    """a function merges all SET log files into one csv file.

    :param fd_path: a folder that stores source csv files
    :type fd_path: the folder path
    """
    if os.path.isdir(fd_path):
        files = glob.glob(os.path.join(fd_path, "*.csv"))
        files.sort()
        big_df = pd.concat((pd.read_csv(f, skiprows=1, engine='python') for f in files), ignore_index=True, sort=False)
        big_df = big_df[pd.notnull(big_df['STATE'])]
        #<~ combine all date and save as xl file
        fn = os.path.join(fd_path, 'all_SET_merged.xlsx')
        big_df.to_excel(fn)

def merge_cs2k_log(fd_path: str):
    if os.path.isdir(fd_path):
        files = glob.glob(os.path.join(fd_path, "*.xlsx"))
        #<~ sort file names based on its created date
        d = {}
        for f in files:
            t = os.path.getmtime(f)
            d.setdefault(t, f)
        od = collections.OrderedDict(sorted(d.items(), key=lambda t: t[0]))

        #<~ open file, read data, add filename col
        df_li = []
        for t, f in od.items():
            df = pd.read_excel(f)
            df['Lv'] = df['L[cd/m^2]']
            df['file_name'] = os.path.split(f)[-1]
            df['modified_time'] = datetime.datetime.fromtimestamp(t).strftime('%Y-%m-%d %H:%M:%S')
            df.drop(columns=['L[cd/m^2]'], inplace=True)
            df_li.append(df)

        #<~ combine all date and save as csv file
        big_df = pd.concat(df_li, ignore_index=True, sort=False)
        big_df.to_csv(os.path.join(fd_path, 'all_cs2k_merged.csv'))