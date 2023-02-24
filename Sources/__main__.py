import agent_initializetion
# from http.client import RemoteDisconnected
# from urllib3.exceptions import ProtocolError
# from Sources.active_directory import *
# from Sources.colvir import *
import pyautogui
from agent_initializetion import *


def main():
    # ! Поиск заявок в BPM
    all_users = []
    success = False
    chrome = None

    while True:
        try:
            chrome = ChromeOtbasy()
            all_users = chrome.bpm_parse()
            success = True
        except Exception as error:
            if isinstance(error, RemoteDisconnected) or isinstance(error, ProtocolError):
                print("Remote end closed connection without response, trying to restart the WebDriver...")
            elif isinstance(error, ValueError):
                if str(error) == 'Заявки не найдены':
                    print(str(error))
            else:
                print("Error occurred:", error)
                break
        finally:
            # chrome.quit()
            if success:
                break

    if len(all_users) < 1:
        print('Нет заявок на создание учётной записи')
        return

    kill_browser()

    for user in all_users:
        # ! Добавление пользователя в Active Directory
        try:
            user = insert_to_ad(user)
        except ValueError:
            continue
        add_group_ad(user)
        sleep(4)

        explorer = None
        success = False

        while True:
            try:
                kill_browser()
                explorer = ExplorerOtbasy()
                explorer.bpm_edit(user)
                success = True
            except Exception as error:
                if isinstance(error, RemoteDisconnected) or isinstance(error, ProtocolError):
                    print("Remote end closed connection without response, trying to restart the WebDriver...")
                elif isinstance(error, ValueError):
                    continue
                else:
                    print("Error occurred:", error)
                    break
            finally:
                sleep(2)
                # explorer.quit()
                if success:
                    break

        # ! Создание почтового ящика (Exchange)
        chrome = None
        success = False
        while True:
            try:
                kill_browser()
                chrome = ChromeOtbasy()
                chrome.exchange_login()
                chrome.exchange_start(user)
                success = True
            except Exception as error:
                if isinstance(error, RemoteDisconnected) or isinstance(error, ProtocolError):
                    print("Remote end closed connection without response, trying to restart the WebDriver...")
                elif isinstance(error, selenium.common.exceptions.InvalidElementStateException):
                    continue
                else:
                    print("Error occurred:", error)
                    break
            finally:
                # chrome.quit()
                if success:
                    break

        # ! Создание учётной записи Skype (Lync)
        explorer = None
        success = False
        while True:
            try:
                kill_browser()
                explorer = ExplorerOtbasy()
                explorer.lync_login()
                explorer.lync_start(user)
                success = True
            except Exception as error:
                if isinstance(error, RemoteDisconnected) or isinstance(error, ProtocolError):
                    print("Remote end closed connection without response, trying to restart the WebDriver...")
                else:
                    print("Error occurred:", error)
                    break
            finally:
                # explorer.quit()
                if success:
                    break

        # ! Добавление пользователя в структуру документолога (при необходимости)
        chrome = None
        success = False
        while True:
            try:
                kill_browser()
                chrome = ChromeOtbasy()
                chrome.doc_start(user)
                chrome.exchange_start(user)
                success = True
            except Exception as error:
                if isinstance(error, RemoteDisconnected) or isinstance(error, ProtocolError):
                    print("Remote end closed connection without response, trying to restart the WebDriver...")
                else:
                    print("Error occurred:", error)
                    break
            finally:
                # chrome.quit()
                if success:
                    break

        # ! Создание учётной записи в Colvir
        if 'АБИС "Colvir"' in user['roles'].keys():
            colvir = Colvir(user)
            colvir.kill()
            colvir.colvir_create_user()
            colvir.colvir_add_permissions()  # TODO: Нет полной информации по действиям робота (опционально)
            colvir.kill()

            # ! Проверка наличия карты
            colvir = Colvir(user)
            colvir.colvir_is_card_exist()
            colvir.kill()

        else:
            print('Не запрашивает Colvir')


if __name__ == '__main__':
    main()
    # sleep(10)
    # print('hello')
    # myScreenshot = pyautogui.screenshot()
    # myScreenshot.save(r'C:\Users\robot.ad\Desktop\111.png')
