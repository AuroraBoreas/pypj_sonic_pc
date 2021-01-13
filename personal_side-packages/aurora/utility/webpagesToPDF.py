#!/usr/bin/env python
# coding: utf-8

# # A code snippet to save webpages as pdf @ZL, 20200507
# ### Why?
# for sake of offline usage
# ### how?
# see the following approach
# ### merit
# pure python library
# ### demerit
# works on this pc, check development environment below pls
# 
# #### DevEnv
# * platform         : Windows
# * platform-release : 10
# * platform-version : 10.0.17763
# * architecture     : AMD64
# * hostname         : SSV619288-1995A
# * ip-address       : 192.168.8.101
# * mac-address      : 30:24:32:50:a6:8f
# * processor        : Intel64 Family 6 Model 142 Stepping 10, GenuineIntel
# * ram              : 7 GB
# * python -V        : 3.7.4 (tags/v3.7.4:e09359112e, Jul  8 2019, 19:29:22) [MSC v.1916 32 bit (Intel)]
# * screensize       : [1280, 720]

# # STEP1: Python web scraping

import requests, pyautogui
from bs4 import BeautifulSoup

def convert_webpage_to_pdf(webpages, chrome_exe_path, PDF_preview_time=8):
    """
    STEP2: using <font color='blue'>pyautogui</font> to 'convert' webpages to pdf via <font color='blue'>chrome</font>
    
    """
    import subprocess, time, pyautogui
    from subprocess import check_call

    for webpage in webpages:
        # open webpage with chrome.exe
        p = subprocess.Popen('"{0}" {1}'.format(chrome_exe_path, webpage))
        time.sleep(5)
        # send print hotkey to call out print interface
        pyautogui.hotkey('ctrl', 'p')
        time.sleep(8)
        # send enter hotkey to confirm
        pyautogui.press('enter')
        time.sleep(3)
        # send enter hotkey to confirm
        pyautogui.press('enter')
        time.sleep(1)
        # send close tab hotkey to close this tab
        pyautogui.hotkey('ctrl', 'w')
    return

if __name__ == "__main__":
    # get html doc from website address via requests
    website = r"https://blog.miguelgrinberg.com/post/the-flask-mega-tutorial-part-i-hello-world"
    html_doc = requests.get(website).text

    # extract data from html doc
    soup = BeautifulSoup(html_doc, 'html.parser')
    webpages = []
    for link in soup.find_all('a'):
        try:
            href = link.get('href')
            content = link.contents[0].replace(':', '')
            link_head = "https://blog.miguelgrinberg.com"
            if '/post/' in href and 'Chapter' in content:
                webpages.append(link_head + href)
        except:
            pass
        
    webpages = webpages[:-1]
    chrome_exe_path = r"D:\\GoogleChromePortable\\App\\Chrome-bin\\chrome.exe"
    preview_time = 8
    convert_webpage_to_pdf(webpages=webpages, chrome_exe_path=chrome_exe_path, PDF_preview_time=preview_time)

