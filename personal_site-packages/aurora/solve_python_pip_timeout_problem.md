---
title: python environment
---
# Python pip timeout issue
I hit a snag when I tried to ```pip install numpy``` due to the Great Wall.
Here is workaround to solve it

## Description of pip timeout problem
```cmd
ReadTimeoutError: HTTPSConnectionPool(host='pypi.python.org', port=443): Read tim
```

## Solution
there are two solutions available as follows.<br>
method 1. manually change source link(personally dislike)<br>

```command
pip install numpy -i https://pypi.doubanio.com/simple/
```

other sources<br>
阿里云 http://mirrors.aliyun.com/pypi/simple/ <br>
中国科技大学 https://pypi.mirrors.ustc.edu.cn/simple/ <br>
豆瓣(douban) http://pypi.douban.com/simple/ <br>
清华大学 https://pypi.tuna.tsinghua.edu.cn/simple/ <br>
中国科学技术大学 http://pypi.mirrors.ustc.edu.cn/simple/ <br>

method 2. change default setting(recommend)

- create or modify setting file <br>
<font color='blue'>linux: ~/.pip/pip.conf</font><br>
<font color='blue'>windows: %HOMEPATH%\pip\pip.ini</font>
<br>

- write or copy following content into the setting file<br>
```[global]
index-url = http://mirrors.aliyun.com/pypi/simple/
[install]
trusted-host=mirrors.aliyun.com
```
