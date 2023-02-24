import ctypes
import json
import os
import pywinauto
import win32clipboard
# from .Robot.Utils import ProcessCommunicator
from pyPythonRPA.Robot import pythonRPA
from pyPythonRPA.Robot.Tools import PythonRPA as Tools
from pyPythonRPA.Robot.UIDesktop import UIOSelector_Get_UIOList, UIOSelector_Get_UIO, UIOSelector_Exist_Bool, UIOSelectorSecs_WaitAppear_Bool, UIOSelectorSecs_WaitDisappear_Bool
from time import sleep

if ctypes.windll.user32.GetKeyboardLayout(0) != 67699721:
    pythonRPA.keyboard.press("shift+alt")


class byDesk:
    @staticmethod
    def mkdir(path: str):
        os.makedirs(path, exist_ok=True)
        return path

    class json:
        @staticmethod
        def read(path):
            with open(path, 'r', encoding='utf-8') as fp:
                data = json.load(fp)
            return data

        @staticmethod
        def write(path, data):
            with open(path, 'w+', encoding='utf-8') as fp:
                json.dump(data, fp, ensure_ascii=False)

    class serialisation:
        def __init__(self, serialisation_path, make_backups=False):
            self.serialisation_path = serialisation_path
            self.make_backups = make_backups
            self.data = None
            self.backup = list()
            self.__start()

        def __start(self):
            if os.path.isfile(self.serialisation_path):
                self.update_in()
            else:
                self.update_out(self.data)

        def __backup(self):
            if self.data and len(self.data) and self.make_backups:
                self.backup.append(self.data)

        def print(self, unpacked=True):
            print('[ SERIALISATION ]')
            if unpacked:
                for each in self.data:
                    print_ = (each, '-', self.data[each]) if self.data.__class__ is dict else self.data
                    print(*print_)
            else:
                print(self.data)

        def update_in(self):
            self.__backup()
            self.data = byDesk.json.read(self.serialisation_path)
            return self.data

        def update_data(self, data, key=None):
            if key:
                self.data[key] = data
            else:
                self.data.append(data)
            return self

        def update_out(self, data):
            byDesk.json.write(self.serialisation_path, data)
            return self.data

    class keyboard:
        @staticmethod
        def write(text, timing_before=0.3, timing_after=0.3, timing_between=0.005):
            Tools.keyboard.write(text, timing_after, timing_before, timing_between)

        @staticmethod
        def press(keys, timing_before=0.3, timing_after=0.3, n_time=1):
            Tools.keyboard.press(keys, n_time, timing_after, timing_before)

    class clipboard:
        @staticmethod
        def get(timing_before=0.3):
            sleep(timing_before)
            win32clipboard.OpenClipboard()
            lResult = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            win32clipboard.CloseClipboard()
            return lResult

        @staticmethod
        def set(text: str, timing_before=0.3, timing_after=0.3):
            def set_(text_):
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text_)
                win32clipboard.CloseClipboard()

            sleep(timing_before)
            set_(text)
            n = 100
            while byDesk.clipboard.get() != text and n:
                n -= 1
                sleep(0.1)
            if n == 0:
                raise ValueError('Cannot set to clipboard')
            sleep(timing_after)

    class application:
        def __init__(self, path):
            self.app = pywinauto.Application()
            self.path = path
            self.process = None

        def start(self, timeout=15):
            self.process = self.app.start(self.path, timeout=timeout).process
            sleep(1)

        def connect(self, **kwargs):
            arg_list = [None, None, None, None]
            if 'handle' in kwargs:
                arg_list[1] = kwargs['handle']
            if 'process' in kwargs:
                arg_list[0] = kwargs['process']
            if 'path' in kwargs:
                arg_list[2] = kwargs['path']
            else:
                arg_list[2] = self.path
            if 'timeout' in kwargs:
                arg_list[3] = kwargs['timeout']

            self.app.connect(process=arg_list[0], handle=arg_list[1], path=arg_list[2], timeout=arg_list[3])

        def quit(self, soft=True):
            self.app.kill(soft=soft)

    def __init__(self, maintitle: dict, programpath=None, debug=False):
        self.main_title = [maintitle]
        self.program_path = programpath
        self.debug = debug
        if programpath:
            self.process = byDesk.application(path=programpath).start()
            self.set_focus()
            self.maximize()

    # Native

    def set_focus(self):
        self.__wait(selector=self.main_title, appear=True, timeout=15)
        UIOSelector_Get_UIOList(self.main_title).set_focus()

    def maximize(self):
        UIOSelector_Get_UIOList(self.main_title).maximize()

    # Extended

    def find_element(self, selector, wait_appear=15, delay_before=0):
        selector_ = self.main_title.copy()
        selector_.append(selector)
        sleep(delay_before)
        if wait_appear:
            self.__wait(selector=selector_, appear=True, timeout=wait_appear)
        deskobject = self.__get_element(selector=selector_, multiple=False)
        el = DeskElement(maintitle=self.main_title, deskobject=deskobject, selector=selector_, debug=self.debug)
        return el

    def find_elements(self, selector, by='xpath', wait_appear=15, delay_before=0):
        selector_ = self.main_title.copy()
        selector_.append(selector)
        sleep(delay_before)
        if wait_appear:
            self.__wait(selector=selector_, appear=True, timeout=wait_appear)
        deskobjects = self.__get_element(selector=selector_, multiple=True)
        els = []
        for each in deskobjects:
            els.append(DeskElement(maintitle=self.main_title, deskobject=each, selector=selector_, debug=self.debug))
        return els

    def wait_element(self, selector, appear=True, timeout=15, delay_before=0):
        selector_ = self.main_title.copy()
        selector_.append(selector)
        sleep(delay_before)
        flag = self.__wait(selector_,  appear, timeout)
        return flag

    # Privat methods

    def __wait(self, selector, appear, timeout):
        wait_ = UIOSelectorSecs_WaitAppear_Bool if appear else UIOSelectorSecs_WaitDisappear_Bool
        flag = wait_(inSpecificationList=selector, inWaitSecs=timeout)
        if self.debug:
            print(selector, "appeared:", flag)
        return flag

    def __get_element(self, selector, multiple):
        if not multiple:
            deskobject = UIOSelector_Get_UIO(inSpecificationList=selector)
        else:
            deskobject = UIOSelector_Get_UIOList(inSpecificationList=selector)
        return deskobject


class DeskElement:
    """DeskElement"""

    def __init__(self, maintitle, deskobject, selector, debug):
        self.main_title = maintitle
        self.deskobject = deskobject
        self.selector = selector
        self.debug = debug

    # Related

    def find_element(self, selector, wait_appear=15, delay_before=0):
        selector_ = self.selector.copy()
        selector_.append(selector)
        sleep(delay_before)
        if wait_appear:
            self.__wait(selector=selector_, appear=True, timeout=wait_appear)
        deskobject = self.__get_element(selector=selector_, multiple=False)
        el = DeskElement(maintitle=self.main_title, deskobject=deskobject, selector=selector_, debug=self.debug)
        return el

    def find_elements(self, selector,wait_appear=15, delay_before=0):
        selector_ = self.selector.copy()
        selector_.append(selector)
        sleep(delay_before)
        if wait_appear:
            self.__wait(selector=selector_, appear=True, timeout=wait_appear)
        deskobjects = self.__get_element(selector=selector_, multiple=True)
        els = []
        for each in deskobjects:
            els.append(DeskElement(maintitle=self.main_title, deskobject=each, selector=selector_, debug=self.debug))
        return els

    def wait_element(self, selector, appear=True, timeout=15, delay_before=0):
        selector_ = self.selector.copy()
        selector_.append(selector)
        sleep(delay_before)
        flag = self.__wait(selector_,  appear, timeout)
        return flag

    # Privat methods

    def __wait(self, selector, appear, timeout):
        wait_ = UIOSelectorSecs_WaitAppear_Bool if appear else UIOSelectorSecs_WaitDisappear_Bool
        flag = wait_(inSpecificationList=selector, inWaitSecs=timeout)
        if self.debug:
            print(selector, "appeared:", flag)
        return flag

    def __get_element(self, selector, multiple):
        if not multiple:
            deskobject = UIOSelector_Get_UIO(inSpecificationList=selector)
        else:
            deskobject = UIOSelector_Get_UIOList(inSpecificationList=selector)
        return deskobject

    # Extended

    def wait_disappear(self, timeout=15, delay_before=0):
        sleep(delay_before)
        return self.__wait(selector=self.selector, appear=False, timeout=timeout)

    # Actions

    def click(self, delay_before=0, delay_after=0):
        sleep(delay_before)
        try:
            self.deskobject.ensure_visible()
        except Exception as e:
            del e
        self.deskobject.click()
        sleep(delay_after)
        return self

    def double_click(self, delay_before=0, delay_after=0):
        sleep(delay_before)
        try:
            self.deskobject.ensure_visible()
        except Exception as e:
            del e
        self.deskobject.double_click()
        sleep(delay_after)
        return self

    def send_keys(self, text: str, delay_before=0, delay_after=0, click_before=False, scroll=True, clear=False):
        sleep(delay_before)
        try:
            self.deskobject.ensure_visible()
        except Exception as e:
            del e
        if click_before:
            self.deskobject.click()
            sleep(0.3)
        if clear:
            self.clear()
        self.deskobject.type_keys(text)
        sleep(delay_after)
        return self

    def get_text(self, delay_before=0, texts=False):
        sleep(delay_before)
        if texts:
            value = " ".join(self.deskobject.texts())
        else:
            value = self.deskobject.get_value()
        return value

    def clear(self, delay_before=0):
        sleep(delay_before)
        try:
            self.deskobject.ensure_visible()
        except Exception as e:
            del e
        self.deskobject.type_keys('^a{DELETE}')
        return self
