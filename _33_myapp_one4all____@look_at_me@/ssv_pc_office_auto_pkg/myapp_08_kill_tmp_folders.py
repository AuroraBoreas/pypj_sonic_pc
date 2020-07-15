#!/usr/bin/env python
# coding: utf-8

# Automatically remove all **.tmp** directories in <font color='blue'>AppData\Local\Temp</font> <br>
# @ZL, 20191213

# In[19]:


import os, shutil, tempfile

def remove_temp_dirs(tmp_fd):
    for root, dirs, files in os.walk(tmp_fd):
        for dir_name in dirs:
            if dir_name.startswith('.') or dir_name.endswith('.tmp'):
                shutil.rmtree(os.path.join(root, dir_name), ignore_errors=True)
        for fn in files:
            if fn.endswith('OProcSessId.dat') or fn.endswith('.TMP') or fn.endswith('.log') or fn.startswith('URL'):
                try:
                    os.remove(os.path.join(root, fn))
                except PermissionError:
                    pass


def main():
    tmp_fd = tempfile.gettempdir() #<~ get path of AppData\Local\Temp
    remove_temp_dirs(tmp_fd)

if __name__ == '__main__':
    main()
