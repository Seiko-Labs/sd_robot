import datetime
import io
from time import sleep
from typing import Optional
import pandas as pd
from selenium.webdriver.support.select import Select
from active_directory import is_user_in_ad
from config import *
import keyboard
import psutil
import pywinauto.findwindows
import selenium.webdriver.ie.options
from pyPythonRPA.Robot import pythonRPA
import pyautogui
from selenium.webdriver.common import keys
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from transliterate import translit
from urllib3.exceptions import MaxRetryError
from webdriver_manager.microsoft import IEDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver import Ie, Chrome, ActionChains
from selenium.webdriver.chrome.service import Service as CService
from selenium.webdriver.ie.service import Service
from selenium.common.exceptions import NoSuchElementException, SessionNotCreatedException
from init import ie_settings
import cv2


def kill_browser():
    processes = ('iexplore.exe', 'IEDriverServer.exe', 'EXCEL.EXE')
    try:
        for proc in psutil.process_iter():
            if proc.name() in processes:
                proc.kill()
    except Exception as e:
        print('Error on closing Internet Explorer: ' + str(e))


class ChromeOtbasy(Chrome):
    def __init__(self, options=None):
        if options:
            self.options = options
        else:
            self.options = selenium.webdriver.ie.options.Options()
            self.options.add_argument("--enable-smooth-scrolling")
            self.options.add_argument("--start-maximized")
            self.options.add_argument("--no-sandbox")
            self.options.add_argument("--disable-dev-shm-usage")
            self.options.add_argument("--disable-extensions")
            self.options.add_argument("--disable-notifications")
            self.options.add_argument("--ignore-ssl-errors=yes")
            self.options.add_argument("--ignore-certificate-errors")
            self.options.add_argument('--allow-insecure-localhost')
        super().__init__(service=CService(ChromeDriverManager().install()), options=options)
        self.implicitly_wait(15)
        self.maximize_window()

    def exchange_login(self):
        self.get('https://exch-02.hcsbkkz.loc/ecp/UsersGroups/NewMailboxOnPremises.aspx?pwmcid=7&ReturnObjectType=1')
        sleep(1)
        self.find_element(By.XPATH, '//input[@name="username"]').clear()
        self.find_element(By.XPATH, '//input[@name="username"]').send_keys(str(f'hcsbkkz\\{login}'))
        self.find_element(By.XPATH, '//input[@name="password"]').clear()
        self.find_element(By.XPATH, '//input[@name="password"]').send_keys(password)
        # sleep(2)
        self.find_element(By.XPATH, '//div[@class="signinbutton"]').send_keys(keys.Keys.ENTER)
        sleep(5)

    def exchange_start(self, data: dict):
        sleep(1)

        # * Find in list
        self.find_element(By.XPATH,
                          '//button[@id="ResultPanePlaceHolder_NewMailbox_contentContainer_pickerUser_ctl00_browseButton"]').send_keys(
            keys.Keys.ENTER
        )
        sleep(2)
        try:
            for handle in self.window_handles:
                self.switch_to.window(handle)
                if self.title == 'Select User - Entire Forest':
                    break
        except MaxRetryError:
            self.exchange_start(data=data)
        if len(self.find_elements(By.XPATH, '//table/tbody/tr')) >= 1:
            self.find_element(By.XPATH, '//input[@class="filterTextBox TextNoClearButton"]').send_keys(data['username'])
            self.find_element(By.XPATH, '//a[@title="Search"]').send_keys(keys.Keys.ENTER)
        sleep(5)
        try:
            self.find_element(By.XPATH, '//*[contains(@class, "ListViewItem")]').send_keys(keys.Keys.ENTER)
        except NoSuchElementException:
            print('Почта уже создана')
            return

        for handle in self.window_handles:
            self.switch_to.window(handle)
            if self.title == 'User Mailbox':
                break
        sleep(1)
        # * Save new mail button
        self.find_element(By.XPATH, '//button[@id="ResultPanePlaceHolder_ButtonsPanel_btnCommit"]').send_keys(
            keys.Keys.ENTER)
        sleep(5)

        # * Find created mail
        self.get('https://exch-02.hcsbkkz.loc/ecp/UsersGroups/Mailboxes.slab?reqId=1662020805816&showhelp=false#')
        sleep(5)
        self.find_element(By.XPATH, "//a[@title='Search']").send_keys(keys.Keys.ENTER)
        self.find_element(By.XPATH, "//input[@title='search']").send_keys(data['username'])
        self.find_element(By.XPATH, "//input[@title='search']").send_keys(keys.Keys.ENTER)
        sleep(5)

        pyautogui.click(50, 250)

        # * Disable limit
        # while True:
        #     if pyautogui.locateCenterOnScreen(
        #             r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\web\disable_all.png'):
        #         pyautogui.click(pyautogui.locateCenterOnScreen(
        #             r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\web\disable_all.png'))
        #         sleep(1)
        #         try:
        #             self.find_element(By.XPATH, "//button[@id='dlgModalError_Yes']").send_keys(keys.Keys.ENTER)
        #         except:
        #             pythonRPA.keyboard.press('ENTER')
        #         sleep(4)
        #     else:
        #         break

        # * Add options
        self.find_element(By.XPATH, "//a[@title='Edit']").send_keys(keys.Keys.ENTER)
        sleep(4)
        for handle in self.window_handles:
            self.switch_to.window(handle)
            if self.title == 'Edit User Mailbox':
                break
        self.find_element(By.XPATH, "//a[@name='MailboxUsage']").send_keys(keys.Keys.ENTER)
        try:
            self.find_element(By.XPATH, "//button[@id='dlgModalError_OK']").send_keys(keys.Keys.ENTER)
        except:
            ...
        self.find_element(By.ID, 'bookmarklink_1').click()

        self.find_element(By.ID,
                          'ResultPanePlaceHolder_Mailbox_MailboxUsage_contentContainer_ctl00_moreOptionsUsage_label').click()
        sleep(1)
        try:
            self.find_element(By.XPATH, '//*[text()="Customize the quota settings for this mailbox"]').click()
        except:
            self.find_element(By.XPATH, '//*[text()="Customize the quota settings for this mailbox"]').send_keys(keys.Keys.ENTER)

        inputs = self.find_elements(By.XPATH, '//input[@role="textfield"]')
        input_values = ['0.19', '0.2', '0.2']
        for input_, value in zip(inputs, input_values):
            input_.clear()
            input_.send_keys(value)
            sleep(1)
        sleep(3)

        self.find_element(By.ID, 'bookmarklink_5').click()
        sleep(3)
        id_list = [
            'ResultPanePlaceHolder_Mailbox_MailboxFeatures_contentContainer_UMMailboxPlaceHolder_disableEAS_button',
            'ResultPanePlaceHolder_Mailbox_MailboxFeatures_contentContainer_UMMailboxPlaceHolder_disableMOWA_button',
            'ResultPanePlaceHolder_Mailbox_MailboxFeatures_contentContainer_disableOWA_button',
            'ResultPanePlaceHolder_Mailbox_MailboxFeatures_contentContainer_ctl07_enableArchivePopup_button'
        ]
        for el in id_list:
            element = self.find_element(By.ID, el)
            if 'Disable' in element.get_attribute('title'):
                element.send_keys(keys.Keys.ENTER)
                sleep(1)
                self.find_element(By.XPATH, "//button[@id='dlgModalError_Yes']").send_keys(keys.Keys.ENTER)
                sleep(3)


        # sleep(5)
        # click = pyautogui.locateCenterOnScreen(
        #     r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\web\radiobutton.png')
        # pyautogui.click(click)
        # sleep(2)
        # inputs = self.find_elements(By.XPATH, "//input[@role='textfield']")
        # locale = ['0.19', '0.2', '0.2']
        # for i, j in zip(locale, inputs):
        #     j.clear()
        #     j.send_keys(i)
        #     sleep(1)
        # sleep(3)

        # * Final save
        self.find_element(By.XPATH, '//button[@id="ResultPanePlaceHolder_ButtonsPanel_btnCommit"]').send_keys(
            keys.Keys.ENTER)

        pythonRPA.keyboard.press('ENTER')
        sleep(4)

    def bpm_parse(self) -> list:
        self.get('http://bpm/site/Matrix/UserMessages.aspx')
        sleep(3)
        pythonRPA.keyboard.write('robot.ad', timing_after=1, timing_before=3)
        pythonRPA.keyboard.press('TAB')
        pythonRPA.keyboard.write('Asd12345678', timing_after=1)
        pythonRPA.keyboard.press('TAB', timing_after=1)
        pythonRPA.keyboard.press('ENTER', timing_after=2)
        main_window = self.current_window_handle
        sleep(5)

        # * Первичный сбор заявок
        users = self.find_elements(By.XPATH, '//table[@id="GV_UserMessages"]/tbody/tr[not(@align)]')
        data = []
        for i in users:
            user = i.find_elements(By.TAG_NAME, 'td')
            data.append({'username': user[5].text, 'department': user[4].text,
                         'link': user[0].find_element(By.TAG_NAME, 'a')})

        # * Удаление заявок на создание учётной записи
        to_remove = []
        for user in data:
            if user['username'] == 'Исмухан Асет Ерланұлы':
                continue
            if is_user_in_ad(user):
                to_remove.append(user)

        for user in to_remove:
            data.remove(user)

        # * Сбор данных в словарь
        for i in data:
            sleep(2)
            self.find_element(By.XPATH, f'//a[text()="{i["username"]}"]/parent::td/parent::tr/td/a').send_keys(
                keys.Keys.ENTER)
            sleep(2)
            self.switch_to.window(self.window_handles[-1])
            i['title'] = self.find_element(By.ID, 'lbl_emp_pos').text  # Должность
            i['ID'] = self.find_element(By.ID, 'LBL_TABID').text  # Табельный номер
            i['manager'] = self.find_element(By.XPATH,
                                             '//label[text()="Руководитель:"]/following::label').text  # Руководитель
            i['description'] = self.find_element(By.ID, 'lbl_employee_name').text  # Регистрационный номер
            i['filial'] = self.find_element(By.XPATH, '//label[text()="Банк:"]/following::label').text  # Филиал
            i['department'] = self.find_element(By.ID, 'lbl_emp_departamentn').text  # Подразделение

            # Сбор ролей в словарь
            role_table = self.find_elements(By.XPATH, '//tr[contains(@id, "maintr")][not(@style="display: none;")]')

            roles = {}
            for role in role_table:
                current_line = role.find_elements(By.TAG_NAME, 'td')
                if current_line[0].text in roles.keys():
                    roles[current_line[0].text].append(current_line[1].text)
                else:
                    roles[current_line[0].text] = [current_line[1].text, ]
            i['roles'] = roles  # Роли
            i = replace_code(i)

            # Управление (подразделение)
            if len(self.find_elements(By.ID, 'lbl_emp_depn')) == 1:
                i['branch'] = self.find_element(By.ID, 'lbl_emp_depn').text

            # Доступ в интернет
            if 'Интернет' not in i['roles'].keys():
                i['roles']['Интернет'] = ['INET_MIN_ACCESS']
            for inet in internet_hierarchy:
                if inet in i['roles']['Интернет']:
                    i['roles']['Интернет'] = [inet]

            # Проверка логина на наличие специальных символов
            i['temp_name'] = i['username']
            for j in i['username']:
                if j in alph_map.keys():
                    i['temp_name'] = i['temp_name'].replace(j, alph_map[j])
            text = translit(i['temp_name'], 'ru', reversed=True).split()
            if is_spec_alphabet(i['username']):
                i['login'] = f"{text[0]}.{text[1][0:2]}"
                i['login'] = i['login'].lower()
                i['user'] = f"{text[0]}_{text[1][0:2]}"
                i['user'] = i['user'].upper()
            else:
                i['login'] = f"{text[0]}.{text[1][0]}"
                i['login'] = i['login'].lower()
                i['user'] = f"{text[0]}_{text[1][0]}"
                i['user'] = i['user'].upper()

            if i['filial'] != 'АО "Жилищный строительный сберегательный банк "Отбасы банк"':
                i['login'] = f"{i['filial_c']}.{i['login']}"
                i['user'] = f"{i['filial_c']}_{i['user']}"

            i['login'] = str(i['login']).replace('\'', '').lower()
            i['user'] = str(i['user']).replace('\'', '').upper()

            i['mail'] = f'{i["login"]}@otbasybank.kz'  # Почта
            i = replace_filial(i)

            for k in i['roles']:
                i['roles'][k] = list(set(i['roles'][k]))

            print(i)
            self.close()
            sleep(1)
            self.switch_to.window(main_window)

        return data

    def doc_start(self, data: dict):
        if 'СЭД «Documentolog»' not in data['roles'].keys():
            return

        # * Authorize
        self.get('https://doc.hcsbk.kz/structure/index')
        self.find_element(By.XPATH, "//input[@id='login']").send_keys(login)
        self.find_element(By.XPATH, "//input[@id='password']").send_keys(password)
        self.find_element(By.XPATH, "//input[@id='submit']").click()
        sleep(1)
        pythonRPA.keyboard.press('ENTER')
        if len(self.find_elements(By.XPATH, "//input[@value='Все равно войти']")) == 1:
            sleep(1)
            pythonRPA.keyboard.press('ENTER')

        # data = replace_filial(data, doc_map)

        if 'ДДО' in data['department'] and 'ОПЕРАТОР' in str(data['department']).upper():
            data['department'] = 'Департамент дистанционного обслуживания'

        # * Choose department
        if data['filial_type'] == 'Пользователи ЦА':
            self.find_element(By.XPATH, '//input[@id="structure_search_input"]').send_keys(data['branch'])
            els = self.find_elements(By.XPATH, '//li[@class="ui-menu-item"]')
            els[-1].click()
            sleep(1)
            # selector = self.find_element(By.XPATH, f'//span[contains(.,"{data["branch"]}")]/parent::*/following::*')
            selector = self.find_element(By.XPATH, f'//p[@class="selected"]/span[@class="buttons"]/span')
            selector.click()
            sleep(3)
            selector1 = self.find_element(By.XPATH,
                                          '//span[@class="drop_down_list"][@style=""]/a[contains(.,"сотрудника")]')
            selector1.click()
            sleep(3)
            menu_items = self.find_elements(By.XPATH, "//ul[@id='tabtitles']/li")
        elif data['department'] in department_exceptions:
            sleep(1)
            self.find_element(By.XPATH, '//input[@id="structure_search_input"]').send_keys(data['department'])
            els = self.find_elements(By.XPATH, f'//li[@class="ui-menu-item"]/a[text()=\'{data["department"]}\']')
            els[-1].click()
            sleep(1)
            selector = self.find_element(By.XPATH, f'//p[@class="selected"]/span[@class="buttons"]/span')
            selector.click()
            sleep(3)
            selector1 = self.find_element(By.XPATH,
                                          '//span[@class="drop_down_list"][@style=""]/a[contains(.,"сотрудника")]')
            selector1.click()
            sleep(3)
            menu_items = self.find_elements(By.XPATH, "//ul[@id='tabtitles']/li")
        elif 'ТРУДОВЫЕ' in str(data['department']).upper() or 'ТРУДОВОЕ' in str(data['department']).upper():
            sleep(1)
            self.find_element(By.XPATH, '//input[@id="structure_search_input"]').send_keys(data['filial_doc'])
            els = self.find_elements(By.XPATH, f'//li[@class="ui-menu-item"]/a[contains(.,\'{data["filial_doc"]}\')]')
            els[-1].click()
            sleep(1)
            selector = self.find_element(By.XPATH, f'//p[@class="selected"]/span[@class="buttons"]/span')
            selector.click()
            sleep(3)
            selector1 = selector.find_element(By.XPATH,
                                              '//span[@class="drop_down_list"][@style=""]/a[contains(.,"сотрудника")]')
            selector1.click()
            sleep(3)
            menu_items = self.find_elements(By.XPATH, "//ul[@id='tabtitles']/li")
        elif data['filial_type'] == 'Пользователи филиалов':
            # data = replace_filial(data, doc_map)
            self.find_element(By.XPATH, '//input[@id="structure_search_input"]').send_keys(data['department'])
            els = self.find_elements(By.XPATH, f'//li[@class="ui-menu-item"][contains(.,"{data["filial_map"]}")]')
            els[-1].click()
            sleep(1)
            selector = self.find_element(By.XPATH, f'//p[@class="selected"]/span[@class="buttons"]/span')
            selector.click()
            sleep(3)
            selector = self.find_element(By.XPATH,
                                         '//span[@class="drop_down_list"][@style=""]/a[contains(.,"сотрудника")]')
            selector.click()
            sleep(3)
            menu_items = self.find_elements(By.XPATH, "//ul[@id='tabtitles']/li")

        elif data['department'] in department_exceptions:
            sleep(1)
            self.find_element(By.XPATH, '//input[@id="structure_search_input"]').send_keys(data['department'])
            els = self.find_elements(By.XPATH, f'//li[@class="ui-menu-item"]/a[text()=\'{data["filial_map"]}\']')
            els[-1].click()
            sleep(1)
            selector = self.find_element(By.XPATH, f'//p[@class="selected"]/span[@class="buttons"]/span')
            selector.click()
            sleep(3)
            selector1 = self.find_element(By.XPATH,
                                          '//span[@class="drop_down_list"][@style=""]/a[contains(.,"сотрудника")]')
            selector1.click()
            sleep(3)
            menu_items = self.find_elements(By.XPATH, "//ul[@id='tabtitles']/li")
        else:
            raise ValueError(f'"{data["filial_type"]}" не существует')

        # ! ~~~~~~~~~~~~~~~~~~~~~~~ Fill employee card ~~~~~~~~~~~~~~~~~~~~~~~~~~

        # * Menu "Должность"
        self.execute_script(menu_items[0].get_attribute('onclick'))
        sleep(1)
        self.find_element(By.ID, 'edit_field_display_name').send_keys(data['title'])
        self.find_element(By.ID, 'edit_field_work_phone').send_keys('-')
        self.find_element(By.ID, 'edit_field_f_email').send_keys(data['mail'])

        # * Menu "Сотрудник"
        self.execute_script(menu_items[1].get_attribute('onclick'))
        sleep(1)
        self.find_element(By.ID, 'edit_field_f_surname').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_f_name').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_f_middle_name').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_employee_name').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_employee_name_kz').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_f_display_short_name').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_f_display_short_name_kz').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_dative_name').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_dative_name_kz').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_instrumental_display_name').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_instrumental_display_name_kz').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_genetive_name').send_keys(data['username'])
        sleep(1)
        self.find_element(By.ID, 'edit_field_genetive_name_kz').send_keys(data['username'])
        sleep(1)

        # * Меню "Пользователь"
        self.execute_script(menu_items[2].get_attribute('onclick'))
        sleep(1)
        self.find_element(By.ID, 'edit_field_login').send_keys(data['login'])
        self.find_element(By.ID, 'edit_field_email').send_keys(data['mail'])
        sleep(1)
        self.find_element(By.ID, 'is_active_true').click()

        # * Save
        self.find_element(By.XPATH, '/html/body/div[4]/div[3]/div/button[1]').click()
        sleep(3)
        sleep(2)

    def find_el(self, xpath, by=By.XPATH):
        try:
            return self.find_element(by=by, value=xpath)
        except NoSuchElementException:
            return None


class ExplorerOtbasy(Ie):
    def __init__(self, options=None):
        if options:
            self.options = options
        else:
            self.options = selenium.webdriver.ie.options.Options()
            self.options.add_argument("--enable-smooth-scrolling")
            self.options.add_argument("--start-maximized")
            self.options.add_argument("--no-sandbox")
            self.options.add_argument("--disable-dev-shm-usage")
            self.options.add_argument("--disable-print-preview")
            self.options.add_argument("--disable-extensions")
            self.options.add_argument("--disable-notifications")
            self.options.add_argument("--ignore-ssl-errors=yes")
            self.options.add_argument("--ignore-certificate-errors")
            self.options.add_argument('--allow-insecure-localhost')
        super().__init__(service=Service(IEDriverManager().install()), options=self.options)
        self.maximize_window()
        self.implicitly_wait(15)
        self.proc = 'IEXPLORE.EXE'

    def bpm_start(self):
        self.get('http://bpm/site/Matrix/UserMessages.aspx')
        sleep(2)
        selector = [{'class_name': 'Alternate Modal Top Most', 'backend': 'uia'},
                    {'class_name': 'Credential Dialog Xaml Host', 'control_type': 'Window',
                     'title': 'Windows Security'}, {'class_name': 'ScrollViewer', 'control_type': 'Pane', 'title': ''},
                    {'class_name': 'TextBox', 'control_type': 'Edit', 'title': 'User name'}]
        pythonRPA.bySelector(selector).wait_appear()
        pythonRPA.bySelector(selector).set_text(login)
        selector = [{'class_name': 'Alternate Modal Top Most', 'backend': 'uia'},
                    {'class_name': 'Credential Dialog Xaml Host', 'control_type': 'Window',
                     'title': 'Windows Security'}, {'class_name': 'ScrollViewer', 'control_type': 'Pane', 'title': ''},
                    {'class_name': 'PasswordBox', 'control_type': 'Edit', 'title': 'Password'}]
        pythonRPA.bySelector(selector).wait_appear()
        pythonRPA.bySelector(selector).set_text(password)
        selector = [{'class_name': 'Alternate Modal Top Most', 'backend': 'uia'},
                    {'class_name': 'Credential Dialog Xaml Host', 'control_type': 'Window',
                     'title': 'Windows Security'}, {'class_name': 'Button', 'control_type': 'Button', 'title': 'OK'}]
        pythonRPA.bySelector(selector).click()

    def bpm_parse(self) -> list:
        """
        Собирает данные с заявок матрицы
        :return: Список состоящих из словарей с данными заявок
        """
        self.bpm_start()
        self.get('http://bpm/site/Matrix/UserMessages.aspx')

        def change_window(title: str, reverse=False):
            for window in self.window_handles:
                sleep(1)
                self.switch_to.window(window)
                if reverse:
                    if title in self.title:
                        continue
                else:
                    if title not in self.title:
                        continue

        users = self.find_elements(By.XPATH, '//table[@id="GV_UserMessages"]/tbody/tr[not(@align)]')
        data = []
        for i in users:
            user = i.find_elements(By.TAG_NAME, 'td')
            data.append({'username': user[5].text, 'department': user[4].text,
                         'link': user[0].find_element(By.TAG_NAME, 'a').get_attribute('href')})

        to_remove = []
        for user in data:
            if is_user_in_ad(user):
                to_remove.append(user)

        for user in to_remove:
            data.remove(user)

        main_window = self.current_window_handle

        for i in data:
            self.execute_script(i['link'])
            sleep(3)
            change_window('Заявка')
            self.maximize_window()
            try:
                i['title'] = self.find_element(By.ID, 'lbl_emp_pos').text  # Должность
            except NoSuchElementException:
                sleep(5)
                self.switch_to.window(self.window_handles[-1])
                sleep(2)
                i['title'] = self.find_element(By.ID, 'lbl_emp_pos').text  # Должность
            i['ID'] = self.find_element(By.ID, 'LBL_TABID').text  # Табельный номер
            i['manager'] = self.find_element(By.XPATH,
                                             '//label[text()="Руководитель:"]/following::label').text  # Руководитель
            i['description'] = self.find_element(By.ID, 'lbl_employee_name').text  # Регистрационный номер
            i['filial'] = self.find_element(By.XPATH, '//label[text()="Банк:"]/following::label').text  # Филиал
            i['department'] = self.find_element(By.ID, 'lbl_emp_departamentn').text  # Подразделение

            # Сбор ролей в словарь
            role_table = self.find_elements(By.XPATH, '//tr[contains(@id, "maintr")][not(@style="display: none;")]')
            roles = {}
            for role in role_table:
                current_line = role.find_elements(By.TAG_NAME, 'td')
                if current_line[0].text in roles.keys():
                    roles[current_line[0].text].append(current_line[1].text)
                else:
                    roles[current_line[0].text] = [current_line[1].text, ]
            i['roles'] = roles  # Роли
            i = replace_code(i)

            # Управление (подразделение)
            if len(self.find_elements(By.ID, 'lbl_emp_depn')) == 1:
                i['branch'] = self.find_element(By.ID, 'lbl_emp_depn').text

            # Доступ в интернет
            if 'Интернет' not in i['roles'].keys():
                i['roles']['Интернет'] = 'INET_MIN_ACCESS'
            else:
                if type(i['roles']['Интернет']) == list:
                    for rrl in i['roles']['Интернет']:
                        if 'inet_with_social' in rrl:
                            i['roles']['Интернет'] = 'INET_WITH_SOCIAL'
                        else:
                            i['roles']['Интернет'] = 'INET_MIN_ACCESS'
                elif 'inet_with_social' in i['roles']['Интернет'].lower():
                    i['roles']['Интернет'] = 'INET_WITH_SOCIAL'
                else:
                    i['roles']['Интернет'] = 'INET_MIN_ACCESS'

            # Проверка логина на наличие специальных символов
            i['temp_name'] = i['username']
            for j in i['username']:
                if j in alph_map.keys():
                    i['temp_name'] = i['temp_name'].replace(j, alph_map[j])
            text = translit(i['temp_name'], 'ru', reversed=True).split()
            if is_spec_alphabet(i['username']):
                i['login'] = f"{text[0]}.{text[1][0:2]}"
                i['login'] = i['login'].lower()
                i['user'] = f"{text[0]}_{text[1][0:2]}"
                i['user'] = i['user'].upper()
            else:
                i['login'] = f"{text[0]}.{text[1][0]}"
                i['login'] = i['login'].lower()
                i['user'] = f"{text[0]}_{text[1][0]}"
                i['user'] = i['user'].upper()

            if i['filial'] != 'АО "Жилищный строительный сберегательный банк "Отбасы банк"':
                i['login'] = f"{i['filial_c']}.{i['login']}"
                i['user'] = f"{i['filial_c']}_{i['user']}"
                i['filial_type'] = 'Пользователи филиалов'
            else:
                i['filial_type'] = 'Пользователи ЦА'

            i['login'] = str(i['login']).replace('\'', '').lower()
            i['user'] = str(i['user']).replace('\'', '').upper()

            i['mail'] = f'{i["login"]}@otbasybank.kz'  # Почта
            i = replace_filial(i)
            print(i)
            self.close()
            sleep(1)
            self.switch_to.window(main_window)
            sleep(5)
        return data

    def bpm_edit(self, data):
        self.bpm_start()
        self.get('http://bpm/Site/Administration/Users.aspx')
        self.find_element(By.ID, 'IB_Insert').send_keys(keys.Keys.ENTER)
        sleep(3)
        for handle in self.window_handles:
            self.switch_to.window(handle)
            if 'Адресная книга' in self.title:
                break
        sleep(3)
        self.find_element(By.ID, 'TB_FIO').send_keys(data['username'])
        # sleep(1)
        self.find_element(By.ID, 'ButtonView').send_keys(keys.Keys.ENTER)
        # sleep(1)
        self.find_element(By.XPATH, '//a[contains(.,"Выбрать")]').send_keys(keys.Keys.ENTER)
        selector = pythonRPA.bySelector([{'class_name': 'Alternate Modal Top Most',
                                          'title': 'Адресная книга - Internet Explorer', 'backend': 'uia'},
                                         {'class_name': '#32770', 'control_type': 'Window',
                                          'title': 'Message from webpage'},
                                         {'class_name': '', 'control_type': 'TitleBar', 'title': ''}])
        selector.wait_appear(1)
        if selector.is_exists():
            raise ValueError('Пользователь уже существует')
        else:
            ...

        # * Используется при необходимости внесения ролей в BPM (опционально)
        for handle in self.window_handles:
            self.switch_to.window(handle)
            if 'Адресная книга' not in self.title:
                break
        sleep(5)
        selector = self.find_element(By.ID, 'DDL_Branch')
        # data = replace_filial(data, bpm_map)
        Select(selector).select_by_visible_text(data['filial_bpm'])
        sleep(3)
        selector = self.find_element(By.ID, 'DDL_Dep')
        sleep(1)
        if 'ДДО' in data['department'] or 'Департамент дистанционного обслуживания' in data['department']:
            Select(selector).select_by_visible_text('Департамент дистанционного обслуживания')
        else:
            Select(selector).select_by_index(1)
        sleep(2)
        # self.find_element(By.ID, 'Button_AddRole').send_keys(keys.Keys.ENTER)
        # sleep(3)
        # for handle in self.window_handles:
        #     self.switch_to.window(handle)
        #     if 'ролей' in self.title:
        #         break
        # sleep(5)
        # self.maximize_window()

        # Открытие справочника
        df = pd.read_excel(r'C:\Users\robot.ad\Desktop\RolesChangeRobot\glossary.xlsx', sheet_name='BPM')
        codes = df['Code'].tolist()
        roles = df['Role'].tolist()
        groups = df['Group'].tolist()
        codes_roles_dict = {k: v for k, v in zip(codes, roles)}

        # Добавление группы
        for i in data['roles']['ПО «BPM»']:
            code = i.split(' - ')[0]
            group = None
            for j, k in zip(codes, groups):
                if code == j:
                    group = k
            if str(group) == 'nan' or group is None:
                continue
            self.find_element(By.ID, 'Button_AddExecutor').send_keys(keys.Keys.ENTER)
            sleep(2)
            for handle in self.window_handles:
                self.switch_to.window(handle)
                if 'ролей' in self.title:
                    break
            try:
                selector = self.find_element(By.XPATH, f'//td[text()="{group}"]/preceding-sibling::td/a')
                selector.send_keys(keys.Keys.ENTER)
                self.find_element(By.ID, 'IB_Select').send_keys(keys.Keys.ENTER)
            except NoSuchElementException:
                self.find_element(By.ID, 'IB_CloseWindow').send_keys(keys.Keys.ENTER)
            sleep(2)
            for handle in self.window_handles:
                self.switch_to.window(handle)
                if 'ролей' not in self.title:
                    break

        # Добавление ролей
        current_roles = self.find_elements(By.XPATH, '//select[@id="LB_UserRoles"]/option')
        current_roles = [el.text for el in current_roles]
        if '5-70-12 - Кредитный администратор' in data['roles']['ПО «BPM»'] and '5-70-1 - Кредитный эксперт (роль - Кредитный эксперт)' in data['roles']['ПО «BPM»']:
            data['roles']['ПО «BPM»'].remove('5-70-1 - Кредитный эксперт (роль - Кредитный эксперт)')
        for i in data['roles']['ПО «BPM»']:
            code = i.split(' - ')[0]
            if code in codes_roles_dict:
                role = codes_roles_dict[code]
            else:
                print(f'Роль с кодом {code} не найдена')
                continue
            if role in current_roles or 'Нет роли в BPM' in role:
                continue
            self.find_element(By.ID, 'Button_AddRole').send_keys(keys.Keys.ENTER)
            sleep(2)
            for handle in self.window_handles:
                self.switch_to.window(handle)
                if 'ролей' in self.title:
                    break
            selector = self.find_element(By.XPATH, f'//td[text()="{role}"]/preceding-sibling::td/a')
            selector.send_keys(keys.Keys.ENTER)
            self.find_element(By.ID, 'IB_Select').send_keys(keys.Keys.ENTER)
            sleep(2)
            for handle in self.window_handles:
                self.switch_to.window(handle)
                if 'ролей' not in self.title:
                    break
        sleep(2)
        self.find_element(By.ID, 'Button_OK').send_keys(keys.Keys.ENTER)
        sleep(5)

        # if 'ПО «BPM»' in data['roles'].keys():
        #     if type(data['roles']['ПО «BPM»']) is list:
        #         for i in data['roles']['ПО «BPM»']:
        #             i = i.split(" - ")
        #             if i[1][0].isnumeric():
        #                 role = i[2]
        #             else:
        #                 role = i[1]
        #             if role in bpm_exception.keys():
        #                 role = bpm_exception[role]
        #             self.find_element(By.ID, 'Button_AddRole').send_keys(keys.Keys.ENTER)
        #             sleep(2)
        #             for handle in self.window_handles:
        #                 self.switch_to.window(handle)
        #                 if 'ролей' in self.title:
        #                     break
        #             selector = self.find_element(By.XPATH, f'//td[contains(.,"{role}")]/preceding-sibling::td/a')
        #             selector.send_keys(keys.Keys.ENTER)
        #             self.find_element(By.ID, 'IB_Select').send_keys(keys.Keys.ENTER)
        #             sleep(2)
        #             for handle in self.window_handles:
        #                 self.switch_to.window(handle)
        #                 if 'ролей' not in self.title:
        #                     break
        #             # print(i.split(" - "))
        #     else:
        #         i = data['roles']['ПО «BPM»'].split(" - ")[1]
        #         if i in bpm_exception.keys():
        #             i = bpm_exception[i]
        #         self.find_element(By.ID, 'Button_AddRole').send_keys(keys.Keys.ENTER)
        #         sleep(2)
        #         for handle in self.window_handles:
        #             self.switch_to.window(handle)
        #             if 'ролей' in self.title:
        #                 break
        #         selector = self.find_element(By.XPATH, f'//td[contains(.,"{i}")]/preceding-sibling::td/a')
        #         selector.send_keys(keys.Keys.ENTER)
        #         self.find_element(By.ID, 'IB_Select').send_keys(keys.Keys.ENTER)
        #         sleep(2)
        #         for handle in self.window_handles:
        #             self.switch_to.window(handle)
        #             if 'ролей' not in self.title:
        #                 break
        # else:
        #     selector = self.find_element(By.XPATH, '//td[contains(.,"Участник")]/preceding-sibling::td/a')
        #     # print(selector.get_attribute('outerHTML'))
        #     selector.send_keys(keys.Keys.ENTER)
        # sleep(1)
        # # self.find_element(By.ID, 'IB_Select').send_keys(keys.Keys.ENTER)
        # # sleep(2)
        # for handle in self.window_handles:
        #     self.switch_to.window(handle)
        #     if 'ролей' not in self.title:
        #         break
        # self.find_element(By.ID, 'Button_OK').send_keys(keys.Keys.ENTER)
        # sleep(3)

    def bpm_end(self, data):
        self.bpm_start()
        selector = self.find_element(By.XPATH, f'//a[contains(.,"{data["username"]}")]/parent::td/parent::tr/td/a')
        self.execute_script(selector.get_attribute('href'))
        sleep(5)
        self.switch_to.window(self.window_handles[-1])
        self.maximize_window()
        sleep(3)
        for item in pyautogui.locateAllOnScreen(
                r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\web\checkbox.png'):
            pyautogui.click(item)
        self.find_element(By.ID, 'txtarea_comment').send_keys(
            f"Учётные данные\nЛогин: {data['login']}\nПароль: Asd12345678+\nOutlook: {data['mail']}")
        input()
        self.find_element(By.ID, 'BTN_NEXT').send_keys(keys.Keys.ENTER)

    def exchange_login(self):
        self.get('https://exch-02.hcsbkkz.loc/ecp/UsersGroups/NewMailboxOnPremises.aspx?pwmcid=7&ReturnObjectType=1')
        sleep(1)
        self.find_element(By.XPATH, '//input[@name="username"]').clear()
        self.find_element(By.XPATH, '//input[@name="username"]').send_keys(str(f'hcsbkkz\\{login}'))
        self.find_element(By.XPATH, '//input[@name="password"]').clear()
        self.find_element(By.XPATH, '//input[@name="password"]').send_keys(password)
        # sleep(2)
        self.find_element(By.XPATH, '//div[@class="signinbutton"]').send_keys(keys.Keys.ENTER)
        sleep(5)

    def exchange_start(self, data: dict):
        sleep(1)
        try:
            # * Find in list
            self.find_element(By.XPATH,
                              '//button[@id="ResultPanePlaceHolder_NewMailbox_contentContainer_pickerUser_ctl00_browseButton"]').send_keys(
                keys.Keys.ENTER
            )
            sleep(4)
            try:
                for handle in self.window_handles:
                    self.switch_to.window(handle)
                    if self.title == 'Select User - Entire Forest':
                        break
            except MaxRetryError:
                self.exchange_start(data=data)
            sleep(10)
            if len(self.find_elements(By.XPATH, '//table/tbody/tr')) >= 1:
                self.find_element(By.XPATH, '//input[@class="filterTextBox TextNoClearButton"]').send_keys(data['username'])
                sleep(3)
                self.find_element(By.XPATH, '//a[@title="Search"]').send_keys(keys.Keys.ENTER)
                sleep(15)
            try:
                self.find_element(By.XPATH, '//*[contains(@class, "ListViewItem")]').send_keys(keys.Keys.ENTER)
            except NoSuchElementException:
                print('Почта уже создана')
                return
                # raise ValueError('Exchange: Пользователь не найден')

            for handle in self.window_handles:
                self.switch_to.window(handle)
                if self.title == 'User Mailbox':
                    break
            sleep(1)
            # * Save new mail button
            self.find_element(By.XPATH, '//button[@id="ResultPanePlaceHolder_ButtonsPanel_btnCommit"]').send_keys(
                keys.Keys.ENTER)
            sleep(5)

            # * Find created mail
            self.get('https://exch-02.hcsbkkz.loc/ecp/UsersGroups/Mailboxes.slab?reqId=1662020805816&showhelp=false#')
            sleep(5)
            self.find_element(By.XPATH, "//a[@title='Search']").send_keys(keys.Keys.ENTER)
            self.find_element(By.XPATH, "//input[@title='search']").send_keys(data['username'])
            self.find_element(By.XPATH, "//input[@title='search']").send_keys(keys.Keys.ENTER)
            click = pyautogui.locateCenterOnScreen(
                r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\web\user.png')
            pyautogui.click(click[0], click[1] + 35)
            pyautogui.click(click)
            sleep(5)

            # * Disable limit
            while True:
                if pyautogui.locateCenterOnScreen(
                        r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\web\disable_all.png'):
                    pyautogui.click(pyautogui.locateCenterOnScreen(
                        r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\web\disable_all.png'))
                    sleep(1)
                    try:
                        self.find_element(By.XPATH, "//button[@id='dlgModalError_Yes']").send_keys(keys.Keys.ENTER)
                    except:
                        pythonRPA.keyboard.press('ENTER')
                    sleep(4)
                else:
                    break

            # * Add options
            self.find_element(By.XPATH, "//a[@title='Edit']").send_keys(keys.Keys.ENTER)
            sleep(4)
            for handle in self.window_handles:
                self.switch_to.window(handle)
                if self.title == 'Edit User Mailbox':
                    break
            self.find_element(By.XPATH, "//a[@name='MailboxUsage']").send_keys(keys.Keys.ENTER)
            # sleep(3)
            try:
                self.find_element(By.XPATH, "//button[@id='dlgModalError_OK']").send_keys(keys.Keys.ENTER)
                # sleep(3)
            except:
                ...
            click = pyautogui.locateCenterOnScreen(
                r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\web\moreoptions.png')
            pyautogui.click(click)
            sleep(5)
            click = pyautogui.locateCenterOnScreen(
                r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\web\radiobutton.png')
            pyautogui.click(click)
            sleep(2)
            inputs = self.find_elements(By.XPATH, "//input[@role='textfield']")
            locale = ['0.19', '0.2', '0.2']
            for i, j in zip(locale, inputs):
                j.clear()
                j.send_keys(i)
                sleep(1)
            sleep(3)

            # * Final save
            self.find_element(By.XPATH, '//button[@id="ResultPanePlaceHolder_ButtonsPanel_btnCommit"]').send_keys(
                keys.Keys.ENTER)

            pythonRPA.keyboard.press('ENTER')
            sleep(4)

        except NoSuchElementException as e:
            print(e)
            # sleep(15)

    def lync_login(self):
        self.get('https://lync.hcsbkkz.loc/cscp/#MainFrameViewModel%3DUsers%2CUsers%3DUserSearch')
        sleep(2)
        pythonRPA.keyboard.write(login, timing_before=0.2, timing_between=0.1)
        sleep(1)
        pythonRPA.keyboard.press('TAB')
        pythonRPA.keyboard.write(password, timing_before=0.2, timing_between=0.1)
        sleep(1)
        pythonRPA.keyboard.press('ENTER')
        sleep(8)

    def lync_start(self, data: dict):
        self.get('https://lync.hcsbkkz.loc/cscp/#MainFrameViewModel%3DUsers%2CUsers%3DUserSearch')
        sleep(5)

        pyautogui.click(pyautogui.locateCenterOnScreen(
            r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\web\enableusers.png', confidence=0.5))
        sleep(2)

        click = pyautogui.locateCenterOnScreen(
            r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\web\add_skype.png')
        pyautogui.click(click)
        sleep(5)
        click = pyautogui.locateCenterOnScreen(
            r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\web\search.png')
        pyautogui.click(click)
        sleep(1)

        keyboard.write(data['username'])
        sleep(2)
        pyautogui.click(
            pyautogui.locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\web\find.png'))
        sleep(3)
        pyautogui.click(
            pyautogui.locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\web\ok.png'))
        sleep(3)
        pyautogui.click(
            pyautogui.locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\web\UPN.png'))
        sleep(2)
        pyautogui.click(
            pyautogui.locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\web\assign.png')
        )
        sleep(1)
        pythonRPA.keyboard.write("Lync.hcsbkkz.loc")
        sleep(1)

        pyautogui.click(
            pyautogui.locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\web\enable.png'),
            clicks=3)


if __name__ == '__main__':
    ...
