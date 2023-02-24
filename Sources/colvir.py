import itertools
import os
import time
from time import sleep
import psutil
from config import *
from pyautogui import locateCenterOnScreen, click, locateOnScreen, locateAllOnScreen
from pyPythonRPA.Robot import pythonRPA
from transliterate import translit
from web import *
import cv2
import pytesseract


class Colvir:
    def __init__(self, data: dict):
        self.data = data
        self.proc = ('COLVIR.EXE', 'CBSADM.EXE')
        self.default_wait = 15
        self.instance = '//10.10.2.60/CBS40'
        self.process = pythonRPA.bySelector([{'class_name': 'TfrmCssMenu', 'title': 'Выбор режима', 'backend': 'uia'},
                                             {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                                             {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
                                             {'class_name': 'TExEdit', 'control_type': 'Pane', 'title': ' '},
                                             {'class_name': 'TClMaskEdit', 'control_type': 'Edit', 'title': ''}])

    def colvir_start_adm(self):
        os.startfile(r"C:\CBS_R\CBSADM.exe")

        # Селектор окна входа
        pythonRPA.bySelector([{'class_name': 'TfrmLoginDlg', 'title': 'Вход в систему', 'backend': 'uia'},
                              {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                              {'class_name': 'TEdit', 'control_type': 'Edit', 'title': ''}]).wait_appear(
            120)

        # Авторизация
        pythonRPA.keyboard.write(colvir_adm_login, timing_before=1)
        pythonRPA.keyboard.press('TAB')
        sleep(1)
        pythonRPA.keyboard.write(colvir_adm_password, timing_before=1)
        pythonRPA.keyboard.press('TAB')
        sleep(1)
        pythonRPA.keyboard.write(self.instance)
        pythonRPA.keyboard.press('TAB')
        sleep(1)
        pythonRPA.keyboard.press('ENTER')

    def colvir_start(self):
        os.startfile(r"C:\CBS_R\COLVIR.exe")

        # Селектор окна входа
        pythonRPA.bySelector([{'class_name': 'TfrmLoginDlg', 'title': 'Вход в систему', 'backend': 'uia'},
                              {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                              {'class_name': 'TEdit', 'control_type': 'Edit', 'title': ''}]).wait_appear(120)

        # Авторизация
        pythonRPA.keyboard.write(colvir_login, timing_between=0.1, timing_before=1)
        pythonRPA.keyboard.press('TAB')
        sleep(1)
        pythonRPA.keyboard.write(colvir_password, timing_between=0.1, timing_before=1)
        pythonRPA.keyboard.press('TAB')
        sleep(1)
        pythonRPA.keyboard.press('ENTER')
        selector = pythonRPA.bySelector(
            [{'class_name': 'TfrmCssAppl', 'title': 'Банковская система COLVIR', 'backend': 'uia'},
             {'class_name': '#32770', 'control_type': 'Window', 'title': 'Colvir Banking System'},
             {'class_name': 'CCPushButton', 'control_type': 'Button', 'title': 'OK'}])
        selector.wait_appear(self.default_wait)
        selector.click()
        selector = pythonRPA.bySelector([{'class_name': 'TfrmCssMenu', 'title': 'Выбор режима', 'backend': 'uia'},
                                         {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                                         {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
                                         {'class_name': 'TExEdit', 'control_type': 'Pane', 'title': ' '},
                                         {'class_name': 'TClMaskEdit', 'control_type': 'Edit', 'title': ''}])
        selector.wait_appear(10)
        if not selector.is_exists():
            self.kill()
            self.colvir_start()

    def colvir_is_card_exist(self):
        # self.colvir_start()

        # Селектор окна ввода процесса Colvir
        self.process.wait_appear(self.default_wait)
        self.process.set_text('PRS')  # Название процесса в Colvir
        pythonRPA.keyboard.press('ENTER')

        # Поиск пользователя по ФИО
        selector = pythonRPA.bySelector([{'class_name': 'TfrmFilterParams', 'title': 'PRS_GR4', 'backend': 'uia'},
                                         {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                                         {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                                         {'class_name': 'TExBitBtn', 'control_type': 'Button', 'title': 'OK'}])
        selector.wait_appear(self.default_wait)
        click(locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\clear.png',
                                   confidence=0.7))
        sleep(1)
        pythonRPA.keyboard.press('TAB', 7, timing_after=1)
        pythonRPA.keyboard.write(self.data['username'])
        selector.click()

        # Селектор вывода
        # selector = pythonRPA.bySelector([{'class_name': 'TMessageForm', 'title': 'Подтверждение', 'backend': 'uia'}])
        # selector.wait_appear(self.default_wait)
        # if selector.is_exists():
        #     check = False
        # check = True

        if 'АБИС "Colvir"' in self.data['roles'].keys():
            sleep(1)
            click(locateCenterOnScreen(
                r'C:\Users\robot.ad\Desktop\UserCreationRobot_COPY\Sources\img\colvir\operation.png'))
            sleep(1)
            pythonRPA.keyboard.press('DOWN', timing_after=1)
            pythonRPA.keyboard.press('LEFT', timing_after=1)
            sleep(1)
            pythonRPA.keyboard.press('ENTER', timing_before=0.5)
            pythonRPA.keyboard.press('ENTER', timing_before=0.5)
            selector = pythonRPA.bySelector([{'class_name': 'TMessageForm', 'title': 'Подтверждение', 'backend': 'uia'},
                                             {'class_name': 'TButton', 'control_type': 'Button', 'title': 'Да'}])
            selector.wait_appear(self.default_wait)
            selector.click()
            pythonRPA.keyboard.write(self.data['user'], timing_before=1, timing_between=0.2)
            selector = pythonRPA.bySelector(
                [{'class_name': 'TfrmUserConfigDlg', 'title': 'Пользователь', 'backend': 'uia'},
                 {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                 {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                 {'class_name': 'TExBitBtn', 'control_type': 'Button', 'title': 'OK'}])
            selector.wait_appear(self.default_wait)
            selector.click()
            pythonRPA.keyboard.press('ENTER', timing_before=1)

    def colvir_add_permissions(self):

        def create_position():
            click(locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\create.png'))
            selector = pythonRPA.bySelector([{'class_name': 'TfrmUsrDetail', 'title': 'Позиция ', 'backend': 'uia'},
                                             {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
                                             {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
                                             {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
                                             {'class_name': 'TExDBEdit', 'control_type': 'Pane', 'title': ' '},
                                             {'class_name': 'TDBEdit', 'control_type': 'Edit', 'title': ''}])
            selector.wait_appear(10)
            pythonRPA.keyboard.write(self.data['user'])  # Код позиции
            pythonRPA.keyboard.press('TAB')
            sleep(1)
            pythonRPA.keyboard.write(self.data['username'])  # ФИО
            pythonRPA.keyboard.press('TAB')
            sleep(1)
            df = pd.read_excel(r"C:\Users\robot.ad\Desktop\RolesChangeRobot\glossary.xlsx", sheet_name='AWG')
            codes = df['Code']
            groups = df['Workgroup']
            awg = None

            for role in self.data['roles']['АБИС "Colvir"']:
                role = role.split(' - ')[0]
                for code, group in zip(codes, groups):
                    if role in code:
                        awg = group
                    if role == '40-4' and self.data['title'] == 'Администратор зала':
                        awg = '!!!АРМ Администратор зала'
            if awg is None:
                print('Рабочая группа не определена')
                return

            pythonRPA.keyboard.write(awg)  # Рабочее место
            pythonRPA.keyboard.press('ENTER')
            sleep(1)
            selector = pythonRPA.bySelector([{'class_name': 'TfrmUsrDetail', 'title': 'Позиция 1', 'backend': 'uia'},
                                             {'class_name': '#32770', 'control_type': 'Window',
                                              'title': 'Colvir Banking System'},
                                             {'class_name': '', 'control_type': 'TitleBar', 'title': ''}])
            if selector.is_exists():
                raise ValueError(f"Группа \"{get_group(self.data)}\" не существует")
            pythonRPA.keyboard.press('TAB')
            sleep(1)
            if '5-70-105 - Оператор видеобанкинга' in self.data['roles']['ПО «BPM»']:
                self.data['filial_colvir'] = '31'
            pythonRPA.keyboard.write(self.data['filial_colvir'])  # Код подразделения
            pythonRPA.keyboard.press('ENTER')

            # Кнопка сохранить
            click(locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\save.png'))

        def link_position():
            sleep(2)
            click(locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\linkusers.png'))
            pythonRPA.bySelector(
                [{'class_name': 'TfrmStfUsrLst', 'title': 'Связи пользователей с позицией ', 'backend': 'uia'},
                 {'class_name': '', 'control_type': 'TitleBar', 'title': ''}]).wait_appear(self.default_wait)
            click(
                locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\create_link.png',
                                     confidence=0.7))
            pythonRPA.bySelector([{'class_name': 'TfrmStfUsrDtl', 'title': 'Связь c позицией ', 'backend': 'uia'},
                                  {'class_name': '', 'control_type': 'TitleBar', 'title': ''}]).wait_appear(
                self.default_wait)
            pythonRPA.keyboard.write(self.data['user'], timing_before=0.1)
            pythonRPA.keyboard.press('ENTER')
            pythonRPA.keyboard.write('010119000000', timing_between=0.1)
            pythonRPA.keyboard.press('ENTER')
            pythonRPA.keyboard.write('010120500000', timing_between=0.1)
            click(locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\save.png'))
            sleep(2)

        def create_user_grant():
            def create_authority(title, points_list, authority_code, n_times):
                pythonRPA.bySelector([{'class_name': 'TfrmHieFrmOrg', 'backend': 'uia'},
                                      {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
                                      {'class_name': 'TPanel', 'control_type': 'Pane', 'title': 'pnlGrid'},
                                      {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                                      {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''}]).click()
                sleep(1)
                click(locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\get.png'))
                sleep(1)
                pythonRPA.keyboard.write('Документы', timing_after=1, timing_between=0.1)
                pythonRPA.keyboard.press('ENTER', n_time=2, timing_after=1)
                sleep(1)
                click(
                    locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\create.png'))
                sleep(1)
                pythonRPA.keyboard.write(authority_code, timing_after=1, timing_between=0.1)
                pythonRPA.keyboard.press('TAB')
                pythonRPA.keyboard.write(f'{title}_{self.data["user"]}', timing_after=1, timing_between=0.1)
                sleep(1)
                click(locateCenterOnScreen(
                    r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\permissions_step4.png',
                    confidence=0.9))
                sleep(1)

                # Доступные действия
                check_points = locateAllOnScreen(
                    r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\cpoint.png')

                for i, point in enumerate(check_points):
                    if i in points_list:
                        click(point)
                    else:
                        click(point, clicks=2)
                    sleep(1)
                pythonRPA.keyboard.press('PAGE_DOWN')  # Сохранение дубля

                # Создание правил для полномочия
                sleep(2)
                pythonRPA.keyboard.write('1')
                pythonRPA.keyboard.press('TAB', n_times, 0.5)
                pythonRPA.keyboard.write(self.data['user'])
                sleep(1)
                pythonRPA.keyboard.press('PAGE_DOWN')  # Сохранение
                sleep(1)

                # Закрытие окон создания документов
                selector = pythonRPA.bySelector(
                    [{'class_name': 'TfrmGrantMonytDtl', 'title': 'Полномочие', 'backend': 'uia'},
                     {'class_name': '', 'control_type': 'TitleBar', 'title': ''},
                     {'class_name': '', 'control_type': 'Button', 'title': 'Close'}])
                selector.wait_appear(10)
                selector.click()
                sleep(1)
                pythonRPA.bySelector([{'class_name': 'TfrmHieFrmOrg',
                                       'title': f'Зависимости - Пользователи - {self.data["username"]}',
                                       'backend': 'uia'},
                                      {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
                                      {'class_name': 'TPanel', 'control_type': 'Pane', 'title': 'pnlGrid'},
                                      {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                                      {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''}]).click()
                sleep(1)

            sleep(1)
            df = pd.read_excel(r'C:\Users\robot.ad\Desktop\RolesChangeRobot\glossary.xlsx', sheet_name='USRGRANT')
            codes = df['Code']
            roles = df['Role']
            departments = df['Department']

            roles_to_append = []

            for i in self.data['roles']['АБИС "Colvir"']:
                i = i.split(' - ')[0]
                for code, role, department in itertools.zip_longest(codes, roles, departments):
                    if i == code:
                        if department == 'Any':
                            roles_to_append.append(role)
                        else:
                            if self.data['filial'] == department:
                                role = str(role).split(' | ')
                                for k in role:
                                    roles_to_append.append(k)

            self.process.wait_appear(self.default_wait)
            self.process.set_text('USRGRANT')  # Название процесса в Colvir
            pythonRPA.keyboard.press('ENTER')

            # Селектор кнопки "ОК" фильтра
            selector = pythonRPA.bySelector(
                [{'class_name': 'TfrmFilterParams', 'title': 'Фильтр', 'backend': 'uia'},
                 {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                 {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                 {'class_name': 'TExBitBtn', 'control_type': 'Button', 'title': 'OK'}])
            selector.wait_appear(self.default_wait)
            sleep(1)

            # Очистка фильтра
            click(locateCenterOnScreen(
                r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\clear_pos.png',
                confidence=0.7))
            sleep(1)
            pythonRPA.keyboard.press('TAB', timing_after=1)
            pythonRPA.keyboard.write(self.data['username'], timing_between=0.1)  # Логин пользователя
            selector.click()

            # Настройка полномочий
            click(
                locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\settings.png',
                                     confidence=0.7))
            sleep(4)
            for role in roles_to_append:
                # Назначенные профили
                click(
                    locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\profiles.png',
                                         confidence=0.9))
                sleep(1)

                # Селектор кнопки развернуть окно
                pythonRPA.bySelector([{'class_name': 'TfrmHieFrmOrg', 'backend': 'uia'},
                                      {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
                                      {'class_name': 'TPanel', 'control_type': 'Pane', 'title': 'pnlGrid'},
                                      {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                                      {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''}]).click()
                sleep(1)

                # Поиск профилей
                click(locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\get.png'))
                sleep(1)
                pythonRPA.keyboard.write(f"{role}")
                pythonRPA.keyboard.press('ENTER')
                sleep(2)

                # Добавление найденного профиля
                click(locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\add.png',
                                           confidence=0.7))
                sleep(3)
                pythonRPA.bySelector([{'class_name': 'TfrmHieFrmOrg',
                                       'title': f'Зависимости - Пользователи - {self.data["username"]}',
                                       'backend': 'uia'},
                                      {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
                                      {'class_name': 'TPanel', 'control_type': 'Pane', 'title': 'pnlGrid'},
                                      {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                                      {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''}]).click()
                sleep(1)

            # Нахождение учётной записи для создания дубликатов
            if '40-4 профиль Смена пароля' in roles_to_append:
                return
            click(locateCenterOnScreen(
                r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\permissions_step6.png',
                confidence=0.9))
            sleep(1)
            create_authority(title='Документы', authority_code='T_ORD$', points_list=[0, 1, 4], n_times=9)
            sleep(1)
            create_authority(title='Доступ_к_журнальным_записям', authority_code='T_JRN$', points_list=[0, 1],
                             n_times=4)

            # Добавление созданных полномочий
            types = ['Документы', 'Для разделения доступа к журнальным записям']
            authority_codes = ['Документы', 'Доступ_к_журнальным_записям']

            click(locateCenterOnScreen(
                r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\permissions_step6.png',
                confidence=0.9))
            sleep(1)

            for type_, i in zip(types, authority_codes):
                pythonRPA.bySelector([{'class_name': 'TfrmHieFrmOrg', 'backend': 'uia'},
                                      {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
                                      {'class_name': 'TPanel', 'control_type': 'Pane', 'title': 'pnlGrid'},
                                      {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                                      {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''}]).click()
                sleep(1)
                click(locateCenterOnScreen(
                    r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\permissions_step6.png',
                    confidence=0.9))
                sleep(1)
                click(locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\get.png',
                                           confidence=0.9))
                sleep(1)
                pythonRPA.keyboard.write(type_, timing_after=1)
                pythonRPA.keyboard.press('TAB', timing_after=1)
                pythonRPA.keyboard.write(f"{i}_{self.data['user']}")
                pythonRPA.keyboard.press('ENTER', timing_after=1)
                click(locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\add.png',
                                           confidence=0.9))
                sleep(2)
                pythonRPA.bySelector([{'class_name': 'TfrmHieFrmOrg',
                                       'title': f'Зависимости - Пользователи - {self.data["username"]}',
                                       'backend': 'uia'},
                                      {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
                                      {'class_name': 'TPanel', 'control_type': 'Pane', 'title': 'pnlGrid'},
                                      {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                                      {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''}]).click()
                sleep(1)

            pythonRPA.bySelector(
                [{'class_name': 'TfrmHieFrmOrg', 'title': f'Зависимости - Пользователи - {self.data["username"]}',
                  'backend': 'uia'}, {'class_name': '', 'control_type': 'TitleBar', 'title': ''},
                 {'title': 'Close', 'control_type': 'Button', 'class_name': ''}]).click()
            sleep(1)
            pythonRPA.keyboard.press('ENTER')
            sleep(1)
            pythonRPA.keyboard.press('ENTER')

        def unlink_additional_position():
            self.process.wait_appear(10)
            self.process.set_text('USER')
            pythonRPA.keyboard.press('ENTER')

            selector = pythonRPA.bySelector([{'class_name': 'TfrmFilterParams', 'title': 'Фильтр', 'backend': 'uia'},
                                             {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                                             {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                                             {'class_name': 'TExBitBtn', 'control_type': 'Button',
                                              'title': 'OK'}]).wait_appear(self.default_wait)
            pythonRPA.keyboard.press('TAB', timing_before=1)
            pythonRPA.keyboard.write(self.data['username'], timing_before=1, timing_between=0.3)
            sleep(1)
            # selector.click()
            pythonRPA.keyboard.press('ENTER', n_time=6, timing_after=0.5)
            sleep(1)
            pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract\tesseract.exe"
            image = pyautogui.screenshot(region=(271, 944, 151, 16))

            # OCR text extraction
            text = pytesseract.image_to_string(image, lang='eng+rus')

            if not text[0].isnumeric():
                print('Ошибка в отвязке позиции. Код пользователя не найден')
            else:
                self.process.wait_appear(self.default_wait)
                self.process.set_text('MUSERL')
                click(locateCenterOnScreen(
                    r'C:\Users\robot.ad\Desktop\UserCreationRobot_COPY\Sources\img\colvir\clear.png'))
                selector = pythonRPA.bySelector(
                    [{'class_name': 'TfrmFilterParams', 'title': 'Фильтр', 'backend': 'uia'},
                     {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                     {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                     {'class_name': 'TExBitBtn', 'control_type': 'Button', 'title': 'OK'}])
                selector.wait_appear(self.default_wait)
                pythonRPA.keyboard.write(text, timing_before=1, timing_between=0.1, timing_after=1)
                selector.click()
                pythonRPA.bySelector([{'class_name': 'TfrmUsrList', 'title': 'Позиции', 'backend': 'uia'},
                                      {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
                                      {'class_name': 'TExDBGrid', 'control_type': 'Pane', 'title': ''}]).wait_appear(
                    self.default_wait)
                sleep(2)
                click(locateCenterOnScreen(
                    r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\linkusers.png'))
                pythonRPA.bySelector(
                    [{'class_name': 'TfrmStfUsrLst', 'title': 'Связи пользователей с позицией ', 'backend': 'uia'},
                     {'class_name': '', 'control_type': 'TitleBar', 'title': ''}]).wait_appear(self.default_wait)
                pythonRPA.keyboard.press('F4')
                pythonRPA.bySelector(
                    [{'class_name': 'TfrmStfUsrDtl', 'title': f'Связь c позицией {text}', 'backend': 'uia'},
                     {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
                     {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
                     {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''}]).wait_appear(self.default_wait)
                pythonRPA.bySelector(
                    [{'class_name': 'TfrmStfUsrDtl', 'title': f'Связь c позицией {text}', 'backend': 'uia'},
                     {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
                     {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
                     {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
                     {'class_name': 'TExDBCheckBox', 'control_type': 'CheckBox', 'title': 'Основной владелец'}]).click()
                sleep(2)
                click(locateCenterOnScreen(
                    r'C:\Users\robot.ad\Desktop\UserCreationRobot_COPY\Sources\img\colvir\save.png', confidence=0.9))
                sleep(1)

        self.colvir_start()
        # self.data = replace_filial(self.data, colvir_map)

        # Селектор окна ввода процесса Colvir
        self.process.wait_appear(10)
        self.process.set_text('MUSERL')  # Название процесса в Colvir
        pythonRPA.keyboard.press('ENTER')

        # Селектор окна фильтра
        selector = pythonRPA.bySelector([{'class_name': 'TfrmFilterParams', 'title': 'Фильтр', 'backend': 'uia'},
                                         {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                                         {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                                         {'class_name': 'TBitBtn', 'control_type': 'Button', 'title': 'Пустой'}])
        selector.wait_appear(10)
        selector.click()
        sleep(1)

        # * Создание позиции
        create_position()

        # * Связь пользователя с позицией
        link_position()

        # * Назначение профилей и полномочий
        create_user_grant()

        # * Привязка логина к карточке
        self.colvir_is_card_exist()

        # * Отвязка дополнительной позиции
        unlink_additional_position()

    def is_user_exist(self, num, selector):
        click(locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\group.png'))
        if pythonRPA.bySelector([{'class_name': 'TMessageForm', 'title': 'Информация', 'backend': 'uia'}]).is_exists():
            pythonRPA.bySelector([{'class_name': 'TMessageForm', 'title': 'Информация', 'backend': 'uia'},
                                  {'class_name': 'TButton', 'control_type': 'Button', 'title': 'OK'}]).click()
            num += 1
            self.is_user_exist(num, selector)
        else:
            return

    def colvir_create_user(self):
        def change_pwd():
            click(
                locateCenterOnScreen(
                    r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\password_change.png',
                    confidence=0.9))
            sleep(2)
            pythonRPA.keyboard.write('qwe_12345678', timing_after=1)
            pythonRPA.keyboard.press('TAB', timing_after=1)
            pythonRPA.keyboard.write('qwe_12345678', timing_after=1)
            pythonRPA.keyboard.press('TAB', timing_after=1)
            pythonRPA.keyboard.write('Из-за отсутствия пользователя', 1)
            pythonRPA.keyboard.press('TAB', 2, timing_after=1)
            pythonRPA.bySelector(
                [{'class_name': 'TfrmPswDialogAdm', 'title': f'Смена пароля для {self.data["user"]}', 'backend': 'uia'},
                 {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                 {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                 {'class_name': 'TExBitBtn', 'control_type': 'Button', 'title': 'OK'}]).click()

            pythonRPA.keyboard.press('ENTER')

        sleep(3)

        self.colvir_start_adm()

        # * Клик "Пользователи"
        pythonRPA.bySelector([{'class_name': 'TfrmCssApplAdm',
                               'title': 'Банковская система COLVIR: Администратор безопасности ROBOT_ADM - ROBOT_ADM',
                               'backend': 'uia'},
                              {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''}]).wait_appear(
            10)
        click(locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\users.png'))

        # * Создание пользователя
        selector = pythonRPA.bySelector([{'class_name': 'TfrmFilterParams', 'title': 'Фильтр', 'backend': 'uia'},
                                         {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                                         {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''},
                                         {'class_name': 'TBitBtn', 'control_type': 'Button', 'title': 'Пустой'}])
        selector.wait_appear(3)
        selector.click()

        pythonRPA.bySelector([{'class_name': 'TfrmAdmUsrList', 'title': 'Пользователи', 'backend': 'uia'},
                              {'class_name': 'TPanel', 'control_type': 'Pane', 'title': ''}]).wait_appear(10)
        click(locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\create.png'))

        selector = pythonRPA.bySelector(
            [{'class_name': 'TfrmAdmUsrDetail', 'title': 'Пользователь ', 'backend': 'uia'},
             {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
             {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
             {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
             {'class_name': 'TExDBEdit', 'control_type': 'Pane', 'title': ' '},
             {'class_name': 'TDBEdit', 'control_type': 'Edit', 'title': ''}])
        selector.wait_appear(10)
        selector.click()

        # selector = pythonRPA.bySelector([{'class_name': 'TfrmAdmUsrDetail', 'backend': 'uia'},
        #                                  {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
        #                                  {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
        #                                  {'class_name': 'TExPanel', 'control_type': 'Pane', 'title': ''},
        #                                  {'class_name': 'TExPageControl', 'control_type': 'Tab', 'title': ''},
        #                                  {'class_name': 'TTabSheet', 'control_type': 'Pane',
        #                                   'title': 'Основные данные'},
        #                                  {'class_name': 'TcxGrid', 'control_type': 'Pane', 'title': ''},
        #                                  {'class_name': 'TcxGridSite', 'control_type': 'Pane', 'title': ''}])
        # selector.wait_appear(10)
        sleep(1)
        pythonRPA.keyboard.write(self.data['user'])
        pythonRPA.keyboard.press('ENTER')
        sleep(1)
        pythonRPA.keyboard.write('ALLUSERS_ALL')
        pythonRPA.keyboard.press('ENTER')
        click(locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\code.png'))
        pythonRPA.keyboard.press('ENTER')
        sleep(1)
        pythonRPA.keyboard.write(self.data['user'])
        pythonRPA.keyboard.press('TAB')
        sleep(1)
        pythonRPA.keyboard.write(self.data['username'])
        sleep(1)
        pythonRPA.keyboard.press('TAB')
        pythonRPA.keyboard.write(self.data['username'])
        sleep(1)
        pythonRPA.keyboard.press('TAB', 2, 0.5)
        pythonRPA.keyboard.write(1)
        sleep(1)
        pythonRPA.keyboard.press('TAB')
        pythonRPA.keyboard.write(self.data['user'])
        sleep(1)
        pythonRPA.keyboard.press('TAB', 2, 0.5)
        pythonRPA.keyboard.press('SPACE', 2, 0.5)

        # Сохранить изменения
        click(locateCenterOnScreen(r'C:\Users\robot.ad\Desktop\UserCreationRobot\Sources\img\colvir\save.png'))
        sleep(5)
        change_pwd()
        self.kill()

    def kill(self):
        try:
            for proc in psutil.process_iter():
                if proc.name() in self.proc:
                    proc.kill()
        except Exception as e:
            print('Error on closing Colvir or Colvir admin: ' + str(e))


if __name__ == '__main__':
    ...
