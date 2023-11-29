import os
import traceback
from threading import Thread
from time import sleep

from pyautogui import moveTo

from config import odines_username_rpa, odines_password_rpa
from config import process_list_path
from tools.app import App
from tools.exceptions import ApplicationException, RobotException, BusinessException
from tools.process import kill_process_list


class Odines(App):
    def __init__(self, timeout=60, debug=False, logger=None):
        path_ = r'C:\Program Files\1cv8\common\1cestart.exe'
        super(Odines, self).__init__(path_, timeout=timeout, debug=debug, logger=logger)
        self.keys.CLEAR = self.keys.CLEAN
        self.fuckn_tooltip_selector = {
            "class_name": "V8ConfirmationWindow", "control_type": "ToolTip",
            "visible_only": True, "enabled_only": True, "found_index": 0,
            "parent": self.root
        }
        self.root_selector = {
            "title_re": "1С:Предприятие - Алматы центр / ТОО \"Magnum Cash&Carry\" / Алматы  управление / .*",
            "class_name": "V8TopLevelFrame", "control_type": "Window",
            "visible_only": True, "enabled_only": True, "found_index": 0,
            "parent": None
        }
        self.close_1c_config_flag = False
        Thread(target=self.close_1c_config, daemon=True).start()

    def run(self) -> None:
        self.quit()
        os.system(f'start "" "{self.path.__str__()}"')

        # * launcher ---------------------------------------------------------------------------------------------------
        self.root = self.find_element({
            "title": "Запуск 1С:Предприятия", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": None
        })
        self.find_element({
            "title": "go_copy", "class_name": "", "control_type": "ListItem",
            "visible_only": True, "enabled_only": True, "found_index": 0
        }).click(double=True)
        sleep(3)

        # * authentificator --------------------------------------------------------------------------------------------
        self.root = self.find_element({
            "title": "Доступ к информационной базе", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
            "found_index": 0, "parent": None
        }, timeout=30)
        self.find_element({
            "title": "", "class_name": "", "control_type": "ComboBox", "visible_only": True, "enabled_only": True,
            "found_index": 0
        }).type_keys(odines_username_rpa, self.keys.TAB, click=True, clear=True, protect_first=True)
        self.find_element({
            "title": "", "class_name": "", "control_type": "Edit", "visible_only": True, "enabled_only": True,
            "found_index": 0
        }).set_text(odines_password_rpa)
        self.find_element({
            "title": "OK", "class_name": "", "control_type": "Button", "visible_only": True, "enabled_only": True,
            "found_index": 0
        }).click()

        # * set root window --------------------------------------------------------------------------------------------
        self.root = self.find_element(self.root_selector, timeout=180)

        # * close startup banners --------------------------------------------------------------------------------------
        self.close_all_inner(nav_close_all=True)

        # * close 1c config popup thread flag --------------------------------------------------------------------------
        self.close_1c_config_flag = True

    def quit(self):
        # * close 1c config popup thread flag --------------------------------------------------------------------------
        self.close_1c_config_flag = False

        if self.root:
            # * закрыть окна
            # with suppress(Exception):
            self.close_all_inner(nav_close_all=True)

            # * выход
            # with suppress(Exception):
            self.navigate('Файл', 'Выход')
            if self.wait_element({
                "title": "Завершить работу с программой?", "class_name": "", "control_type": "Pane",
                "visible_only": True, "enabled_only": True, "found_index": 0, "parent": self.root
            }, timeout=5):
                self.find_element({
                    "title": "Да", "class_name": "", "control_type": "Button", "visible_only": True,
                    "enabled_only": True, "found_index": 0, "parent": self.root
                }, timeout=1).click()
                self.wait_element({
                    "title": "Да", "class_name": "", "control_type": "Button", "visible_only": True,
                    "enabled_only": True, "found_index": 0, "parent": self.root
                }, timeout=5, until=False)

        kill_process_list(process_list_path)
        sleep(3)

    def wait_fuckn_tooltip(self):
        # with suppress(Exception):
        if self.root:
            window = self.root
            position = window.element.element_info.rectangle.mid_point()
            moveTo(position[0], position[1])
            self.wait_element(self.fuckn_tooltip_selector, until=False)
            sleep(0.5)

    def close_1c_config(self):
        while True:
            if self.close_1c_config_flag:
                # with suppress(Exception):
                if self.wait_element({
                    "title_re": "В конфигурацию ИБ внесены изменения.*", "class_name": "", "control_type": "Pane",
                    "visible_only": True, "enabled_only": True, "found_index": 0, "parent": self.root
                }, timeout=0):
                    self.find_element({
                        "title": "Нет", "class_name": "", "control_type": "Button",
                        "visible_only": True, "enabled_only": True, "found_index": 0, "parent": self.root
                    }, timeout=0).click()
            sleep(0.5)

    def navigate(self, *steps, maximize_innder=False):
        sleep(1)
        # self.wait_fuckn_tooltip()
        for n, step in enumerate(steps):
            if n:
                if not self.wait_element({
                    "title": step, "class_name": "", "control_type": "MenuItem",
                    "visible_only": True, "enabled_only": True, "found_index": 0, "parent": self.root
                }, timeout=2):
                    if n - 1:
                        self.find_element({
                            "title": steps[n - 1], "class_name": "", "control_type": "MenuItem",
                            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": self.root
                        }, timeout=5).click()
                    else:
                        self.find_element({
                            "title": steps[n - 1], "class_name": "", "control_type": "Button",
                            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": self.root
                        }, timeout=5).click()
                self.find_element({
                    "title": step, "class_name": "", "control_type": "MenuItem",
                    "visible_only": True, "enabled_only": True, "found_index": 0, "parent": self.root
                }, timeout=5).click()
            else:
                self.find_element({
                    "title": step, "class_name": "", "control_type": "Button",
                    "visible_only": True, "enabled_only": True, "found_index": 0, "parent": self.root
                }, timeout=5).click()
        if maximize_innder:
            self.maximize_inner()

    def close_all_inner(self, iter_count=10, manual_close_until=1, nav_close_all=False):
        # * закрыть все внутренние окна через меню
        if nav_close_all:
            close_buttons = self.find_elements({
                "title": "Закрыть", "class_name": "", "control_type": "Button",
                "visible_only": True, "enabled_only": True, "parent": self.root
            }, timeout=1)
            if len(close_buttons) > 1:
                # with suppress(Exception):
                self.close_1c_error()
                # with suppress(Exception):
                self.navigate('Окна', 'Закрыть все')
                # with suppress(Exception):
                self.close_1c_error()

        while True:
            iter_count -= 1
            close_buttons = self.find_elements({
                "title": "Закрыть", "class_name": "", "control_type": "Button",
                "visible_only": True, "enabled_only": True, "parent": self.root
            }, timeout=1)
            if len(close_buttons) > manual_close_until:
                # with suppress(Exception):
                self.close_1c_error()
                # with suppress(Exception):
                self.find_element({
                    "title": "Закрыть", "class_name": "", "control_type": "Button", "visible_only": True,
                    "enabled_only": True, "found_index": len(close_buttons) - 1, "parent": self.root
                }, timeout=1).click()
                # with suppress(Exception):
                self.close_1c_error()
            else:
                break
            if iter_count < 0:
                raise Exception('Не все окна закрыты')

    def maximize_inner(self, timeout=0.5):
        self.root.type_keys('%+r', set_focus=True)
        if self.wait_element({
            "title": "Развернуть", "class_name": "", "control_type": "Button",
            "visible_only": True, "enabled_only": True, "parent": self.root
        }, timeout=timeout):
            self.find_elements({
                "title": "Развернуть", "class_name": "", "control_type": "Button",
                "visible_only": True, "enabled_only": True, "parent": self.root
            })[-1].click()

    def check_1c_error(self, function_name, data=None, count=1):
        root_window = self.root
        while count > 0:
            count -= 1
            # * Конфигурация базы данных не соответствует сохраненной конфигурации -------------------------------------
            if self.wait_element({
                "title": "Конфигурация базы данных не соответствует сохраненной конфигурации.\nПродолжить?",
                "class_name": "", "control_type": "Pane",
                "visible_only": True, "enabled_only": True, "found_index": 0, "parent": None
            }, timeout=0.2):
                error_message = "Конфигурация базы данных не соответствует сохраненной конфигурации"
                raise ApplicationException(error_message, function_name, data)

            # * Строка не найдена --------------------------------------------------------------------------------------
            if self.wait_element({
                "title": "Строка не найдена!", "class_name": "", "control_type": "Pane",
                "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
            }, timeout=0.2):
                error_message = "Строка не найдена"
                raise ApplicationException(error_message, function_name, data)

            # * critical Разрыв соединения -----------------------------------------------------------------------------
            if self.wait_element({
                "title_re": "^.*Удаленный хост принудительно разорвал существующее подключение.*",
                "class_name": "", "control_type": "Pane",
                "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
            }, timeout=0.2):
                error_message = "critical Ошибка разрыв соединения"
                raise ApplicationException(error_message, function_name, data)

            # * critical Ошибка исполнения отчета ----------------------------------------------------------------------
            if self.wait_element({
                "title": "Ошибка исполнения отчета", "class_name": "", "control_type": "Pane",
                "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
            }, timeout=0.2):
                error_message = "critical Ошибка исполнения отчета"
                raise ApplicationException(error_message, function_name, data)

            # * Ошибка при вызове метода контекста ---------------------------------------------------------------------
            if self.wait_element({
                "title_re": "Ошибка при вызове метода контекста (.*)",
                "class_name": "", "control_type": "Pane",
                "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
            }, timeout=0.2):
                error_message = "Ошибка при вызове метода контекста"
                raise ApplicationException(error_message, function_name, data)

            # * Конфликт блокировок при выполнении транзакции ----------------------------------------------------------
            if self.wait_element({
                "title_re": "Конфликт блокировок при выполнении транзакции:.*",
                "class_name": "", "control_type": "Pane",
                "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
            }, timeout=0.2):
                error_message = "Конфликт блокировок при выполнении транзакции"
                raise ApplicationException(error_message, function_name, data)

            # * Операция не выполнена ----------------------------------------------------------------------------------
            if self.wait_element({
                "title": "Операция не выполнена", "class_name": "", "control_type": "Pane",
                "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
            }, timeout=0.2):
                error_message = "Операция не выполнена"
                raise RobotException(error_message, function_name, data)

            # * Введенные данные не отображены в списке, так как не соответствуют отбору -------------------------------
            if self.wait_element({
                "title": "Введенные данные не отображены в списке, так как не соответствуют отбору.",
                "class_name": "", "control_type": "Pane",
                "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
            }, timeout=0.2):
                error_message = "Введенные данные не отображены в списке, так как не соответствуют отбору"
                raise RobotException(error_message, function_name, data)

            # * critical В поле введены некорректные данные ------------------------------------------------------------
            if self.wait_element({
                "title_re": "В поле введены некорректные данные.*", "class_name": "", "control_type": "Pane",
                "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
            }, timeout=0.2):
                error_message = "critical В поле введены некорректные данные"
                raise RobotException(error_message, function_name, data)

            # * critical Не удалось провести ---------------------------------------------------------------------------
            if self.wait_element({
                "title_re": "Не удалось провести.*", "class_name": "", "control_type": "Pane",
                "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
            }, timeout=0.2):
                error_message = "critical Не удалось провести"
                raise BusinessException(error_message, function_name, data)

            # * critical Сеанс работы завершен администратором ---------------------------------------------------------
            if self.wait_element({
                "title": "Сеанс работы завершен администратором.", "class_name": "", "control_type": "Pane",
                "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
            }, timeout=0.2):
                error_message = "critical Сеанс работы завершен администратором"
                raise ApplicationException(error_message, function_name, data)

            # * Сеанс отсутствует или удален ---------------------------------------------------------------------------
            if self.wait_element({
                "title_re": "Сеанс отсутствует или удален.*", "class_name": "", "control_type": "Pane",
                "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
            }, timeout=0.2):
                error_message = "critical Сеанс отсутствует или удален"
                raise ApplicationException(error_message, function_name, data)

            # * critical Неизвестное окно ошибки -----------------------------------------------------------------------
            if self.wait_element({
                "title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
                "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
            }, timeout=0.2):
                error_message = "critical Неизвестное окно ошибки"
                raise RobotException(error_message, function_name, data)

    def close_1c_error(self):
        root_window = self.root
        # * Конфигурация базы данных не соответствует сохраненной конфигурации -----------------------------------------
        message_ = {
            "title": "Конфигурация базы данных не соответствует сохраненной конфигурации.\nПродолжить?",
            "class_name": "", "control_type": "Pane",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": None
        }
        button_ = {
            "title": "Да", "class_name": "", "control_type": "Button",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": None
        }
        if self.wait_element(message_, timeout=0.1):
            self.find_element(button_, timeout=1).click(double=True)
            self.wait_element(message_, timeout=5, until=False)

        # * Строка не найдена ------------------------------------------------------------------------------------------
        message_ = {
            "title": "Строка не найдена!", "class_name": "", "control_type": "Pane",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        button_ = {
            "title": "OK", "class_name": "", "control_type": "Button",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        if self.wait_element(message_, timeout=0.1):
            self.find_element(button_, timeout=1).click(double=True)
            self.wait_element(message_, timeout=5, until=False)

        # * Ошибка исполнения отчета -----------------------------------------------------------------------------------
        message_ = {
            "title_re": "^.*Удаленный хост принудительно разорвал существующее подключение.*",
            "class_name": "", "control_type": "Pane",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        button_ = {
            "title": "OK", "class_name": "", "control_type": "Button",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        if self.wait_element(message_, timeout=0.1):
            self.find_element(button_, timeout=1).click(double=True)
            self.wait_element(message_, timeout=5, until=False)

        # * Ошибка исполнения отчета -----------------------------------------------------------------------------------
        message_ = {
            "title": "Ошибка исполнения отчета", "class_name": "", "control_type": "Pane",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        button_ = {
            "title": "OK", "class_name": "", "control_type": "Button",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        if self.wait_element(message_, timeout=0.1):
            self.find_element(button_, timeout=1).click(double=True)
            self.wait_element(message_, timeout=5, until=False)

        # * Ошибка при вызове метода контекста -------------------------------------------------------------------------
        message_ = {
            "title_re": "Ошибка при вызове метода контекста (.*)", "class_name": "",
            "control_type": "Pane", "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        button_ = {
            "title": "OK", "class_name": "", "control_type": "Button",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        if self.wait_element(message_, timeout=0.1):
            self.find_element(button_, timeout=1).click(double=True)
            self.wait_element(message_, timeout=5, until=False)

        # * Завершить работу с программой? -----------------------------------------------------------------------------
        message_ = {
            "title": "Завершить работу с программой?", "class_name": "", "control_type": "Pane",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        button_ = {
            "title": "Да", "class_name": "", "control_type": "Button",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        if self.wait_element(message_, timeout=0.1):
            self.find_element(button_, timeout=1).click(double=True)
            self.wait_element(message_, timeout=5, until=False)

        # * Операция не выполнена --------------------------------------------------------------------------------------
        message_ = {
            "title": "Операция не выполнена", "class_name": "", "control_type": "Pane",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        button_ = {
            "title": "OK", "class_name": "", "control_type": "Button",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        if self.wait_element(message_, timeout=0.1):
            self.find_element(button_, timeout=1).click(double=True)
            self.wait_element(message_, timeout=5, until=False)

        # * Конфликт блокировок при выполнении транзакции --------------------------------------------------------------
        message_ = {
            "title_re": "Конфликт блокировок при выполнении транзакции:.*", "class_name": "", "control_type": "Pane",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        button_ = {
            "title": "OK", "class_name": "", "control_type": "Button",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        if self.wait_element(message_, timeout=0.1):
            self.find_element(button_, timeout=1).click(double=True)
            self.wait_element(message_, timeout=5, until=False)

        # * Введенные данные не отображены в списке, так как не соответствуют отбору -----------------------------------
        message_ = {
            "title": "Введенные данные не отображены в списке, так как не соответствуют отбору.",
            "class_name": "", "control_type": "Pane",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        button_ = {
            "title": "OK", "class_name": "", "control_type": "Button",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        if self.wait_element(message_, timeout=0.1):
            self.find_element(button_, timeout=1).click(double=True)
            self.wait_element(message_, timeout=5, until=False)

        # * Данные были изменены. Сохранить изменения? -----------------------------------------------------------------
        message_ = {
            "title": "Данные были изменены. Сохранить изменения?", "class_name": "", "control_type": "Pane",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        button_ = {
            "title": "Нет", "class_name": "", "control_type": "Button",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        if self.wait_element(message_, timeout=0.1):
            self.find_element(button_, timeout=1).click(double=True)
            self.wait_element(message_, timeout=5, until=False)

        # * critical В поле введены некорректные данные ----------------------------------------------------------------
        message_ = {
            "title_re": "В поле введены некорректные данные.*", "class_name": "", "control_type": "Pane",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        button_ = {
            "title": "Да", "class_name": "", "control_type": "Button",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        if self.wait_element(message_, timeout=0.1):
            self.find_element(button_, timeout=1).click(double=True)
            self.wait_element(message_, timeout=5, until=False)

        # * critical Не удалось провести -------------------------------------------------------------------------------
        message_ = {
            "title_re": "Не удалось провести \".*", "class_name": "", "control_type": "Pane",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        button_ = {
            "title": "OK", "class_name": "", "control_type": "Button",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        if self.wait_element(message_, timeout=0.1):
            self.find_element(button_, timeout=1).click(double=True)
            self.wait_element(message_, timeout=5, until=False)

        # * Сеанс работы завершен администратором ----------------------------------------------------------------------
        message_ = {
            "title": "Сеанс работы завершен администратором.", "class_name": "", "control_type": "Pane",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        button_ = {
            "title": "Завершить работу", "class_name": "", "control_type": "Button", "visible_only": True,
            "enabled_only": True, "found_index": 0, "parent": root_window
        }
        if self.wait_element(message_, timeout=0.1):
            self.find_element(button_, timeout=1).click(double=True)
            self.wait_element(message_, timeout=5, until=False)

        # * Сеанс отсутствует или удален -------------------------------------------------------------------------------
        message_ = {
            "title_re": "Сеанс отсутствует или удален.*", "class_name": "", "control_type": "Pane",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        button_ = {
            "title": "Завершить работу", "class_name": "", "control_type": "Button",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        if self.wait_element(message_, timeout=0.1):
            self.find_element(button_, timeout=1).click(double=True)
            self.wait_element(message_, timeout=5, until=False)

        # * Выбранное действие не было выполнено -----------------------------------------------------------------------
        message_ = {
            "title": "Выбранное действие не было выполнено! Продолжить?", "class_name": "", "control_type": "Pane",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        button_ = {
            "title": "Да", "class_name": "", "control_type": "Button",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        if self.wait_element(message_, timeout=0.1):
            self.find_element(button_, timeout=1).click(double=True)
            self.wait_element(message_, timeout=5, until=False)

        # * Неизвестное окно ошибки ------------------------------------------------------------------------------------
        selector_ = {
            "title": "1С:Предприятие", "class_name": "V8NewLocalFrameBaseWnd", "control_type": "Window",
            "visible_only": True, "enabled_only": True, "found_index": 0, "parent": root_window
        }
        if self.wait_element(selector_, timeout=0.1):
            self.find_element(selector_).close()

    def approve(self, doc_name: str, function_name: str, try_count: int = 30, delay: float = 0) -> str:
        while True:
            # * нажать Провести
            self.find_element({
                "title": "Провести", "class_name": "", "control_type": "Button",
                "visible_only": True, "enabled_only": True, "found_index": 0
            }, timeout=1).click()

            # * дождаться проведения
            done = self.wait_element({
                "title": "Отмена проведения", "class_name": "", "control_type": "Button",
                "visible_only": True, "enabled_only": True, "found_index": 0
            }, timeout=10)

            # ! выход
            try_count -= 1
            if try_count < 0:
                raise Exception(f'Документ {str(doc_name)} не проведен либо нет номера')

            # ! проверка и закрытие ошибок
            if not done:
                try:
                    self.check_1c_error(function_name)
                except Exception as err:
                    traceback.print_exc()
                    if 'critical' in str(err):
                        raise err
                    self.close_1c_error()
            else:
                doc_num = str(self.find_element({
                    "title": "", "class_name": "", "control_type": "Edit",
                    "visible_only": True, "enabled_only": True, "found_index": 0
                }, timeout=0.1).element.iface_value.CurrentValue.replace(' ', ''))
                if not len(doc_num):
                    continue
                break
            sleep(delay)
        self.check_1c_error(function_name)
        return doc_num

    def deprove(self, doc_name: str, function_name: str, try_count: int = 30, delay: float = 0) -> str:
        while True:
            # * нажать Отмена проведения
            self.find_element({
                "title": "Отмена проведения", "class_name": "", "control_type": "Button",
                "visible_only": True, "enabled_only": True, "found_index": 0
            }, timeout=1).click()

            # * дождаться отмены проведения
            done = self.wait_element({
                "title": "Отмена проведения", "class_name": "", "control_type": "Button",
                "visible_only": True, "enabled_only": False, "found_index": 0
            }, timeout=10)

            # ! выход
            try_count -= 1
            if try_count < 0:
                raise Exception(f'Документ {str(doc_name)}. Не отменяется проведение')

            # ! проверка и закрытие ошибок
            if not done:
                try:
                    self.check_1c_error(function_name)
                except Exception as err:
                    traceback.print_exc()
                    if 'critical' in str(err):
                        raise err
                    self.close_1c_error()
            else:
                doc_num = str(self.find_element({
                    "title": "", "class_name": "", "control_type": "Edit",
                    "visible_only": True, "enabled_only": True, "found_index": 0
                }, timeout=0.1).element.iface_value.CurrentValue.replace(' ', ''))
                if not len(doc_num):
                    continue
                break
            sleep(delay)
        self.check_1c_error(function_name)
        return doc_num
