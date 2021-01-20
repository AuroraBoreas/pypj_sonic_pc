---
title: pyinstaller tips
output: markdown
---

# An assembly of pyinstaller tips
pyinstaller is good, but there is something that may trip you down.

Here are tips when I hit a snag.

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
    ```command
    pyinstaller --debug all my_script.py
    ```

## 03 Pyinstaller builds onefile and noconsole exe
    ```command
    pyinstaller --onefile --noconsole --icon="my_script.ico" my_script.py
    ```


## 99 rectify 3rd-party module (pip install python-pptx)
    **python-pptx**
        bug     : newline in textbox
        solution: .add_textbox, 把python-pptx库源码"../python3.7/site-packages/pptx/oxml/shapes/autoshape.py"第434行注释掉
        ```
        "    <a:p/>\n"
        ```