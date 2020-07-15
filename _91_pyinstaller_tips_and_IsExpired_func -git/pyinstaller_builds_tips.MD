
## 00 Verify users and expired time

    ```python code
    import getpass, datetime
    authorized_users = ['5106001995', '5106001916']
    user = getpass.getuser()

    def is_outdated(file_path, days=1):
        """check if file modification time is outdate"""
        mtime = pathlib.Path(file_path).stat().st_mtime
        mtime = datetime.datetime.fromtimestamp(mtime)
        now = datetime.datetime.now()
        return (now - mtime) > datetime.timedelta(days=days)

    def is_expired(expired_date='Dec 1 2020 8:00AM'):
        """check if current date is expired"""
        ed = datetime.datetime.strptime(expired_date,'%b %d %Y %I:%M%p')
        now = datetime.datetime.now()
        return (now - ed) > datetime.timedelta(days=1)

    if is_expired() or (not user in users):
        pass
    else:
        #Do something
    ```

## 01 pyinstaller “Failed to execute script pyi_rth_pkgres” and missing packages

    ```Python package command
    OS: Win10
    Python: 3.7
    pyinstaller installed by pip install pyinstaller
    fix by (same solution with above, by no need download):

    pip uninstall pyinstaller
    pip install https://github.com/pyinstaller/pyinstaller/archive/develop.zip
    ```

## 02 pyinstaller enters debug mode and to check if anything wrong before onefile package
    ```
    pyinstaller --debug all my_script.py
    ```

## 03 Pyinstaller builds onefile and noconsole exe
    ```
    pyinstaller --onefile --noconsole --icon="my_script.ico" my_script.py
    ```


## 99 Trivial, 3rd-party module (pip install python-pptx)
 @ https://www.jianshu.com/p/7a1e2166d18d
 bug: newline in textbox
 solution: 1,针对.add_textbox
把python-pptx库源码"../python2.7/site-packages/pptx/oxml/shapes/autoshape.py"第434行注释掉
