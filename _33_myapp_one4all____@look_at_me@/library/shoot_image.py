import pyautogui

# shoot image for specific area
def shoot_img(str_img_name, x1, y1, x2, y2):
    """
    take screenshot -> convert to binary data -> put on clipboard
    """
    w = x2 - x1
    h = y2 - y1
    pyautogui.screenshot(str_img_name, region=(x1, y1, w, h))

if __name__ == "__main__":
    x1, y1, x2, y2 = (443, 276, 1779, 1032)
    shoot_img("test.png", x1, y1, x2, y2)
