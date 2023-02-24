import ctypes
import subprocess
from pyPythonRPA.Robot import pythonRPA
from _desktools import byDesk
from time import sleep
import os


def ie_settings():
    subprocess.Popen(r'"C:\Program Files\Internet Explorer\iexplore.exe" www.google.com')
    security_info_window_ok = pythonRPA.bySelector(
        [{"title": "Security Alert", "class_name": "#32770", "backend": "win32"}, {"title": "OK"}])
    security_info_window_ok.wait_appear(20)
    if security_info_window_ok.is_exists():
        security_info_window_dontshow = pythonRPA.bySelector(
            [{"title": "Security Alert", "class_name": "#32770", "backend": "win32"},
             {"title": "&In the future, do not show this warning"}])
        security_info_window_dontshow.click()
        sleep(0.2)
        security_info_window_ok.click()
        sleep(0.2)
    ie = byDesk({"title": "Google - Internet Explorer", "class_name": "IEFrame", "backend": "uia"})
    ie.find_element({"depth_start": 4, "depth_end": 4, "title": "Tools", "control_type": "Button"}).click()
    pythonRPA.time.delay(2)
    pythonRPA.keyboard.press("down", 10)
    pythonRPA.time.delay(2)
    pythonRPA.keyboard.press("enter")
    pythonRPA.time.delay(2)
    pythonRPA.keyboard.press("ctrl+tab")
    pythonRPA.time.delay(2)
    ie_options = byDesk({"title": "Internet Options", "class_name": "#32770", "backend": "win32"})
    checkbox = ie_options.find_element(
        {"title": "Enable &Protected Mode (requires restarting Internet Explorer)"}).deskobject
    if checkbox.is_checked():
        checkbox.click()
    sleep(0.2)
    pythonRPA.keyboard.press("enter")
    sleep(0.2)
    pythonRPA.keyboard.press("enter")
    os.system("taskkill /f /im iexplore.exe")


def check_keyboard():
    if ctypes.windll.user32.GetKeyboardLayout(0) != 67699721:
        pythonRPA.keyboard.press("shift+alt")


if __name__ == '__main__':
    ...
