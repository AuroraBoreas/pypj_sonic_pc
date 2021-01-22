import datetime, os

def generate_new_file_name(old_file_name):
    current_weeknum = '{:02d}'.format(datetime.date.today().isocalendar()[1])
    fractions = old_file_name.split("_WK")
    new_name = fractions[0] + "_WK" + current_weeknum + fractions[-1][2:]
    return  new_name

def change_file_names():
    path = "."
    for file in os.listdir(path):
        if os.path.isfile(file) and (file.startswith('FY20 SMLD') or file.startswith('SSV LCM')):
            new_filename = generate_new_file_name(file)
            os.rename(file, new_filename)
    return

if __name__ == '__main__':
    change_file_names()