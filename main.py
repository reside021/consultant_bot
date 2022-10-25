import os
from datetime import datetime

from selenium_work import open_browser

import tkinter as tk
import tkinter.filedialog as fd
import Pmw

import webbrowser
import logging

logger = logging.getLogger("main")
logger.setLevel(logging.INFO)

newpath = 'log_file'
if not os.path.exists(newpath):
    os.makedirs(newpath)

# create the logging file handler
name_file = f"{datetime.now().year}-{datetime.now().month}-{datetime.now().day}-{datetime.now().hour}-{datetime.now().minute}-{datetime.now().second}"
fh = logging.FileHandler(f"log_file/{name_file}.log")

formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s',)
fh.setFormatter(formatter)

# add handler to logger object
logger.addHandler(fh)


def open_chrome_driver_page():
    webbrowser.open('https://chromedriver.chromium.org/downloads', new=2)


def isfloat(value):
    try:
        float(value)
        return True
    except ValueError:
        return False


def interface_open():
    logger.info("Start the program")
    app = App()
    app.mainloop()


class App(tk.Tk):
    def __init__(self):
        try:
            super().__init__()
            self.title("Консультант бот")
            w = 500  # width for the Tk root
            h = 700  # height for the Tk root

            # get screen width and height
            ws = self.winfo_screenwidth()  # width of the screen
            hs = self.winfo_screenheight()  # height of the screen

            # calculate x and y coordinates for the Tk root window
            x = (ws / 2) - (w / 2)
            y = (hs / 2) - (h / 2)

            # set the dimensions of the screen
            # and where it is placed
            self.geometry('%dx%d+%d+%d' % (w, h, x, y))
            # self.minsize(w, h)
            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(0, weight=1)

            self.global_directory = ""

            # блок кнопок
            fr_buttons = tk.Frame(self)
            btn_start = tk.Button(fr_buttons, text="Запуск бота", command=self.collect_data_for_selenium)
            btn_check_driver = tk.Button(fr_buttons, text="Страница ChromeDriver", command=open_chrome_driver_page)

            fr_buttons.grid(row=0, column=0)
            fr_buttons.grid_rowconfigure(1, weight=1)
            fr_buttons.grid_columnconfigure(1, weight=1)

            # блок подсказок
            balloon = Pmw.Balloon(fr_buttons)
            text_for_balloon = "Необходимо загрузить версию драйвера максимально схожую с версией браузера и распаковать архив в папку с ботом"
            balloon.bind(btn_check_driver, text_for_balloon)

            # блок ввода долга
            entry_frame = tk.Frame(fr_buttons)
            label_entry = tk.Label(entry_frame, text="Сумма долга, руб (?): ")
            self.entry = tk.Entry(entry_frame, width=15, font="Calibri 14", justify='center')
            label_entry.grid(row=0, column=0, sticky='e')
            self.entry.grid(row=0, column=1)

            # блок подсказок
            balloon_entry = Pmw.Balloon(fr_buttons)
            text_for_balloon_entry = "Допустим ввод целых чисел и дробных чисел через точку"
            balloon_entry.bind(label_entry, text_for_balloon_entry)

            # блок для ввода первого дня
            entry_frm_first_day = tk.Frame(fr_buttons)
            label_first_day = tk.Label(entry_frm_first_day, text="Первый день расчета (по умолчанию равен 1)")
            self.entry_first_day = tk.Entry(entry_frm_first_day, width=10, font="Calibri 14", justify='center')
            label_first_day.grid(row=0, column=0, sticky='ns')
            self.entry_first_day.grid(row=1, column=0, sticky='ns')

            # блок для ввода последнего дня
            entry_frm_last_day = tk.Frame(fr_buttons)
            label_last_day = tk.Label(entry_frm_last_day, text="Последний день расчета: ", width=25)
            self.entry_last_day = tk.Entry(entry_frm_last_day, width=10, font="Calibri 14", justify='center')
            label_last_day.grid(row=0, column=0, sticky='ns')
            self.entry_last_day.grid(row=0, column=1, sticky='e')

            # блок для ввода первого месяца
            entry_frm_first_month = tk.Frame(fr_buttons)
            label_first_month = tk.Label(entry_frm_first_month, text="Первый месяц расчета (по умолчанию равен 1)")
            self.entry_first_month = tk.Entry(entry_frm_first_month, width=10, font="Calibri 14", justify='center')
            label_first_month.grid(row=0, column=0, sticky='ns')
            self.entry_first_month.grid(row=1, column=0, sticky='ns')

            # блок для ввода последнего месяца
            entry_frm_last_month = tk.Frame(fr_buttons)
            label_last_month = tk.Label(entry_frm_last_month, text="Последний месяц расчета: ", width=25)
            self.entry_last_month = tk.Entry(entry_frm_last_month, width=10, font="Calibri 14", justify='center')
            label_last_month.grid(row=0, column=0, sticky='ns')
            self.entry_last_month.grid(row=0, column=1, sticky='ns')

            # блок для ввода первого года
            entry_frm_first_year = tk.Frame(fr_buttons)
            label_first_year = tk.Label(entry_frm_first_year, text="Первый год расчета (по умолчанию равен 2013)")
            self.entry_first_year = tk.Entry(entry_frm_first_year, width=10, font="Calibri 14", justify='center')
            label_first_year.grid(row=0, column=0, sticky='ns')
            self.entry_first_year.grid(row=1, column=0, sticky='ns')

            # блок для ввода последнего года
            entry_frm_last_year = tk.Frame(fr_buttons)
            label_last_year = tk.Label(entry_frm_last_year, text="Последний год расчета: ", width=25)
            self.entry_last_year = tk.Entry(entry_frm_last_year, width=10, font="Calibri 14", justify='center')
            label_last_year.grid(row=0, column=0, sticky='ns')
            self.entry_last_year.grid(row=0, column=1, sticky='ns')

            # блок для последнего дня счета
            self.var_last_day = tk.BooleanVar()
            self.check_box_last_day = tk.Checkbutton(fr_buttons, text="Производить счет до вчерашнего дня",
                                                     variable=self.var_last_day,
                                                     font="Calibri 12")
            self.check_box_last_day.select()

            # блок подсказок
            balloon_cb = Pmw.Balloon(fr_buttons)
            text_for_balloon_cb = "В выключенном состоянии счет будет производиться до указанного ниже срока"
            balloon_cb.bind(self.check_box_last_day, text_for_balloon_cb)

            # блок для визуализации
            self.var_visualization = tk.BooleanVar()
            self.check_box_visualization = tk.Checkbutton(fr_buttons, text="Визуализация процесса",
                                                          variable=self.var_visualization,
                                                          font="Calibri 12", onvalue=1)
            self.check_box_visualization.select()

            # блок подсказок
            balloon_visualization = Pmw.Balloon(fr_buttons)
            text_for_balloon_visualization = "При включенной визуализации будет открываться окно браузера,\nпри выключенной - все будет происходить в фоне"
            balloon_cb.bind(self.check_box_visualization, text_for_balloon_visualization)

            # блок кнопок
            btn_choose_folder = tk.Button(fr_buttons, text="Выбор места сохранения", command=self.choose_directory)
            self.label_choose_folder = tk.Label(fr_buttons, text="Место сохранения результатов")

            # блок ввода долга
            entry_delay = tk.Frame(fr_buttons)
            label_entry_delay = tk.Label(entry_delay, text="Задержка между действиями (?)")
            self.entry_delay = tk.Entry(entry_delay, width=15, font="Calibri 14", justify='center')
            label_entry_delay.grid(row=0, column=0)
            self.entry_delay.grid(row=1, column=0)

            # блок подсказок
            balloon_delay = Pmw.Balloon(fr_buttons)
            text_for_balloon_delay = "По умолчанию - 0.2 секунды\nЗадержка влияет на скорость выполнения операций и их точность\nЕсли встречаются неточности, то следует увеличить время задержки"
            balloon_delay.bind(label_entry_delay, text_for_balloon_delay)

            # сетка
            btn_start.grid(row=0, column=0, sticky="ew", padx=5, pady=[10, 5])
            self.check_box_visualization.grid(row=1, column=0, sticky="ew", pady=[0, 25])
            entry_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=10)
            entry_frm_first_day.grid(row=3, column=0, sticky="ew", pady=5)
            entry_frm_first_month.grid(row=4, column=0, sticky="ew", pady=5)
            entry_frm_first_year.grid(row=5, column=0, sticky="ew", pady=5)
            self.check_box_last_day.grid(row=6, column=0, sticky="ew", pady=[10, 5])
            entry_frm_last_day.grid(row=7, column=0, sticky="ew", pady=5)
            entry_frm_last_month.grid(row=8, column=0, sticky="ew", pady=5)
            entry_frm_last_year.grid(row=9, column=0, sticky="ew", pady=5)
            btn_choose_folder.grid(row=10, column=0, sticky="ew", pady=[20, 5])
            self.label_choose_folder.grid(row=11, column=0, sticky="ew")
            btn_check_driver.grid(row=12, column=0, sticky="ew", pady=5)
            entry_delay.grid(row=13, column=0, sticky="ew", pady=5, padx=50)

            logger.info("The interface is built successfully")

        except:
            logger.exception("An interface construction error occurred: ")

    def choose_directory(self):
        try:
            directory = fd.askdirectory(title="Открыть папку", initialdir="/")
            if directory:
                self.global_directory = directory
                self.label_choose_folder['text'] = directory
                logger.info(f"The directory to save is selected: {directory}")
        except:
            logger.exception("Error selecting the save directory: ")

    def collect_data_for_selenium(self):
        try:
            dict_with_data = {}
            debt_sum = self.entry.get()
            if isfloat(debt_sum):
                self.entry.config({"background": "White"})
                dict_with_data['debt_sum'] = debt_sum
            else:
                self.entry.config({"background": "#FF8E73"})
                return

            first_day = self.entry_first_day.get()
            if first_day:
                if first_day.isdigit() and 0 < int(first_day) < 32:
                    self.entry_first_day.config({"background": "White"})
                else:
                    self.entry_first_day.config({"background": "#FF8E73"})
                    return
            else:
                first_day = 1
            dict_with_data['first_day'] = first_day

            first_month = self.entry_first_month.get()
            if first_month:
                if first_month.isdigit() and 0 < int(first_month) < 13:
                    self.entry_first_month.config({"background": "White"})
                else:
                    self.entry_first_month.config({"background": "#FF8E73"})
                    return
            else:
                first_month = 1
            dict_with_data['first_month'] = first_month

            first_year = self.entry_first_year.get()
            if first_year:
                currentYear = datetime.now().year + 1
                if first_year.isdigit() and 1999 < int(first_year) < currentYear:
                    self.entry_first_year.config({"background": "White"})
                else:
                    self.entry_first_year.config({"background": "#FF8E73"})
                    return
            else:
                first_year = 2013
            dict_with_data['first_year'] = first_year

            dict_with_data['var_last_day'] = True

            # дата окончания счета
            if not self.var_last_day.get():
                dict_with_data['var_last_day'] = False

                last_day = self.entry_last_day.get()
                if last_day:
                    if last_day.isdigit() and 0 < int(last_day) < 32:
                        self.entry_last_day.config({"background": "White"})
                    else:
                        self.entry_last_day.config({"background": "#FF8E73"})
                        return
                else:
                    self.entry_last_day.config({"background": "#FF8E73"})
                    return
                dict_with_data['last_day'] = last_day

                last_month = self.entry_last_month.get()
                if last_month:
                    if last_month.isdigit() and 0 < int(last_month) < 13:
                        self.entry_last_month.config({"background": "White"})
                    else:
                        self.entry_last_month.config({"background": "#FF8E73"})
                        return
                else:
                    self.entry_last_month.config({"background": "#FF8E73"})
                    return
                dict_with_data['last_month'] = last_month

                last_year = self.entry_last_year.get()
                if last_year:
                    currentYear = datetime.now().year + 1
                    if last_year.isdigit() and 1999 < int(last_year) < currentYear:
                        self.entry_last_year.config({"background": "White"})
                    else:
                        self.entry_last_year.config({"background": "#FF8E73"})
                        return
                else:
                    self.entry_last_year.config({"background": "#FF8E73"})
                    return
                dict_with_data['last_year'] = last_year

            if not self.global_directory:
                return

            dict_with_data['directory'] = self.global_directory

            dict_with_data['var_visualization'] = self.var_visualization.get()

            var_delay = self.entry_delay.get()
            if not var_delay:
                var_delay = 0.2
            if isfloat(var_delay):
                self.entry.config({"background": "White"})
                dict_with_data['var_delay'] = float(var_delay)
            else:
                self.entry.config({"background": "#FF8E73"})
                return

            logger.info("Successful data collection to start the bot")
            open_browser(dict_with_data)

        except:
            logger.exception("Bot start error: ")


if __name__ == "__main__":
    interface_open()
