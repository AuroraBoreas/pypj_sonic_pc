import pyautogui

def shoot_img(str_img_name, x1, y1, x2, y2):
    """
    # shoot image for specific area
    take screenshot -> convert to binary data -> put on clipboard
    """
    w = x2 - x1
    h = y2 - y1
    pyautogui.screenshot(str_img_name, region=(x1, y1, w, h))