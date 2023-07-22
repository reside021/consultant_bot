import os
from datetime import datetime, date

from selenium_work import open_browser
from tkcalendar import Calendar

import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as fd
import Pmw
import logging


logger = logging.getLogger("main")
logger.setLevel(logging.INFO)

newpath = 'log_file'
if not os.path.exists(newpath):
    os.makedirs(newpath)

# create the logging file handler
name_file = f"{datetime.now().year}-{datetime.now().month}-{datetime.now().day}-{datetime.now().hour}-{datetime.now().minute}-{datetime.now().second}"
fh = logging.FileHandler(f"log_file/{name_file}.log")

formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s', )
fh.setFormatter(formatter)

# add handler to logger object
logger.addHandler(fh)



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
        global text_start_date, text_end_date
        global fr_buttons
        global msg_start_date
        global msg_end_date
        global msg_save_place
        global x,y
        try:
            super().__init__()
            self.title("Консультант бот")
            w = 500  # width for the Tk root
            h = 500  # height for the Tk root

            # get screen width and height
            ws = self.winfo_screenwidth()  # width of the screen
            hs = self.winfo_screenheight()  # height of the screen

            # calculate x and y coordinates for the Tk root window
            x = (ws / 2) - (w / 2)
            y = (hs / 2) - (h / 2)

            # set the dimensions of the screen
            # and where it is placed
            self.geometry('%dx%d+%d+%d' % (w, h, x, y))
            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(0, weight=1)

            self.global_directory = ""



            # блок кнопок
            fr_buttons = tk.Frame()
            btn_start = tk.Button(fr_buttons, text="Запуск бота", command=self.collect_data_for_selenium)

            fr_buttons.grid(row=0, column=0)
            fr_buttons.grid_rowconfigure(1, weight=1)
            fr_buttons.grid_columnconfigure(1, weight=1)

            # блок прогрессбара

            self.progress_bar_frame = tk.Frame(fr_buttons)
            self.progressbar = ttk.Progressbar(self.progress_bar_frame, orient="horizontal", mode="indeterminate", length=400)
            self.progressbar.grid(row=0, column=0, columnspan=2, sticky='ns', padx=10)
            label_progress = tk.Label(self.progress_bar_frame, text="Ведется поиск оптимального драйвера...Ожидайте")
            label_progress.grid(row=1, column=0, columnspan=2, sticky='ns', padx=10)


            # блок ввода долга
            entry_sum_duty = tk.Frame(fr_buttons)
            label_entry = tk.Label(entry_sum_duty, text="Сумма долга, руб (?): ")
            self.entry_sum = tk.Entry(entry_sum_duty, width=15, font="Calibri 14", justify='center')
            label_entry.grid(row=0, column=0, sticky='e')
            self.entry_sum.grid(row=0, column=1)

            # блок подсказок
            balloon_entry = Pmw.Balloon(fr_buttons)
            text_for_balloon_entry = "Допустим ввод целых чисел и дробных чисел через точку"
            balloon_entry.bind(label_entry, text_for_balloon_entry)

            # блок начальной даты
            entry_block_start_date = tk.Frame(fr_buttons)
            self.text_start_date = tk.Label(entry_block_start_date, font=("Calibri", 12), text="ДД.ММ.ГГГГ")
            self.btn_start = tk.Button(entry_block_start_date, text="Выбор даты начала расчета", command=lambda m="start": self.pick_date(m))
            self.msg_start_date = tk.Label(entry_block_start_date, font=("Calibri", 12), foreground='#FF2D00')
            self.btn_start.grid(row=0, column=0, sticky='ns', padx=10)
            self.text_start_date.grid(row=0, column=1, sticky='ns')
            self.msg_start_date.grid(row=1, column=0, sticky='ns', columnspan=2)

            # блок конечной даты
            entry_block_end_date = tk.Frame(fr_buttons)
            self.text_end_date = tk.Label(entry_block_end_date, font=("Calibri", 12), text="ДД.ММ.ГГГГ")
            self.btn_end = tk.Button(entry_block_end_date, text="Выбор даты окончания расчета", command=lambda m="end": self.pick_date(m))
            self.msg_end_date = tk.Label(entry_block_end_date, font=("Calibri", 12), foreground='#FF2D00')
            self.btn_end.grid(row=0, column=0, sticky='ns', padx=10)
            self.text_end_date.grid(row=0, column=1, sticky='ns')
            self.msg_end_date.grid(row=1, column=0, sticky='ns', columnspan=2)

            # блок для последнего дня счета
            self.var_last_day = tk.BooleanVar()
            self.check_box_last_day = tk.Checkbutton(fr_buttons, text="Производить счет до вчерашнего дня",
                                                     variable=self.var_last_day,
                                                     font="Calibri 12")
            self.check_box_last_day.select()

            self.var_begin_month = tk.BooleanVar()
            self.check_box_today_begin_month = tk.Checkbutton(fr_buttons, text="Установить отсчет месяца с выбранного дня",
                                                     variable=self.var_begin_month,
                                                     font="Calibri 12")

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

            # блок выбора места сохранения
            btn_choose_folder = tk.Button(fr_buttons, text="Выбор места сохранения", command=self.choose_directory)
            self.label_choose_folder = tk.Label(fr_buttons, text="Место сохранения результатов")
            self.msg_save_place = tk.Label(fr_buttons, font=("Calibri", 12), foreground='#FF2D00')

            # блок ввода долга
            entry_delay_frame = tk.Frame(fr_buttons)
            label_entry_delay = tk.Label(entry_delay_frame, text="Задержка между действиями (?)")
            self.entry_delay = tk.Entry(entry_delay_frame, width=10, font="Calibri 12", justify='center')
            label_entry_delay.grid(row=0, column=0)
            self.entry_delay.grid(row=0, column=1)

            # блок подсказок
            balloon_delay = Pmw.Balloon(fr_buttons)
            text_for_balloon_delay = "По умолчанию - 0.2 секунды\nЗадержка влияет на скорость выполнения операций и их точность\nЕсли встречаются неточности, то следует увеличить время задержки"
            balloon_delay.bind(label_entry_delay, text_for_balloon_delay)

            # сетка
            btn_start.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
            self.check_box_visualization.grid(row=2, column=0, sticky="ew", pady=5)
            entry_sum_duty.grid(row=3, column=0, sticky="ew", padx=5, pady=5)
            entry_block_start_date.grid(row=4, column=0, sticky="ew", pady=10)
            entry_block_end_date.grid(row=5, column=0, sticky='ew', pady=5)
            self.check_box_last_day.grid(row=6, column=0, sticky="ew", pady=5)
            self.check_box_today_begin_month.grid(row=7, column=0, sticky="ew", pady=5)
            btn_choose_folder.grid(row=8, column=0, sticky="ew", pady=5)
            self.label_choose_folder.grid(row=9, column=0, sticky="ew")
            self.msg_save_place.grid(row=10, column=0, sticky='ew', pady=5)
            entry_delay_frame.grid(row=11, column=0, sticky="ew", pady=5, padx=50)

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
            debt_sum = self.entry_sum.get()
            if isfloat(debt_sum):
                self.entry_sum.config({"background": "White"})
                dict_with_data['debt_sum'] = debt_sum
            else:
                self.entry_sum.config({"background": "#FF8E73"})
                return

            start_period = self.text_start_date.cget("text").split('.')

            if start_period[0].isnumeric():
                self.msg_start_date.config(text="")
            else:
                self.msg_start_date.config(text="Необходимо выбрать начальную дату!")
                return

            first_day = start_period[0]
            dict_with_data['first_day'] = first_day

            first_month = start_period[1]
            dict_with_data['first_month'] = first_month

            first_year = start_period[2]
            dict_with_data['first_year'] = first_year

            dict_with_data['var_last_day'] = True

            # дата окончания счета
            if not self.var_last_day.get():
                dict_with_data['var_last_day'] = False

                end_period = self.text_end_date.cget("text").split('.')

                if end_period[0].isnumeric():
                    self.msg_end_date.config(text="")
                else:
                    self.msg_end_date.config(text="Необходимо выбрать конечную дату!")
                    return

                last_day = end_period[0]
                dict_with_data['last_day'] = last_day

                last_month = end_period[1]
                dict_with_data['last_month'] = last_month

                last_year = end_period[2]
                dict_with_data['last_year'] = last_year

            if self.global_directory:
                self.msg_save_place.config(text="")
            else:
                self.msg_save_place.config(text="Необходимо выбрать место сохранения!")
                return

            dict_with_data['directory'] = self.global_directory

            dict_with_data['var_visualization'] = self.var_visualization.get()

            dict_with_data['var_begin_month'] = self.var_begin_month.get()

            var_delay = self.entry_delay.get()
            if not var_delay:
                var_delay = 0.2
            if isfloat(var_delay):
                self.entry_sum.config({"background": "White"})
                dict_with_data['var_delay'] = float(var_delay)
            else:
                self.entry_sum.config({"background": "#FF8E73"})
                return

            logger.info("Successful data collection to start the bot")
            open_browser(dict_with_data)


        except:
            logger.exception("Bot start error: ")

    def pick_date(self, button_press):
        global cal, date_window

        date_window = tk.Toplevel(fr_buttons)
        date_window.grab_set()
        date_window.title('Выбор даты')
        date_window.geometry('250x220+%d+%d'% (x, y))
        cal = Calendar(date_window, selectmode="day", date_pattern="dd.mm.yyyy")
        cal.place(x=0, y=0)

        submit_btn = tk.Button(date_window, text="Подтвердить", command=lambda m=button_press: self.grab_date(m))
        submit_btn.place(x=80, y=190)

    def grab_date(self, button_press):
        if button_press == "start":
            self.text_start_date.config(text=cal.get_date())
        elif button_press == "end":
            self.text_end_date.config(text=cal.get_date())
        date_window.destroy()

if __name__ == "__main__":
    interface_open()
