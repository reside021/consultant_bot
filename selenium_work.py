import time
import calendar
import logging
from io import StringIO

from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from datetime import datetime
from datetime import date
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from dateutil.relativedelta import *

import pandas as pd
from tkinter import messagebox as mbox

module_logger = logging.getLogger("main.selenium")



def format_date_with_dot(_date):
    _date = _date[:2] + '.' + _date[2:4] + '.' + _date[4:]
    return _date


def open_browser(dict_with_data):
    logger = logging.getLogger("main.selenium.open_browser")

    try:

        debt_sum = float(dict_with_data['debt_sum'])
        first_day = int(dict_with_data['first_day'])
        first_month = int(dict_with_data['first_month'])
        first_year = int(dict_with_data['first_year'])
        var_last_day = dict_with_data['var_last_day']
        directory_for_result = dict_with_data['directory']
        var_visual = dict_with_data['var_visualization']
        var_delay = dict_with_data['var_delay']
        var_begin_month = dict_with_data['var_begin_month']

        if var_last_day:
            last_month = datetime.now().month
            last_year = datetime.now().year
            last_day = datetime.now().day
        else:
            last_day = int(dict_with_data['last_day'])
            last_month = int(dict_with_data['last_month'])
            last_year = int(dict_with_data['last_year'])

        try:
            options = Options()
            if var_visual:
                options.add_argument("--start-maximized")
            else:
                options.add_argument('--headless')
                options.add_argument('--disable-gpu')
            driver = webdriver.Chrome(options=options, service=ChromeService(ChromeDriverManager().install()))
            driver.get("https://calc.consultant.ru/395gk")
        except Exception as ex:
            text = f"Текст ошибки:\n\n{ex}\n\nПроверьте версии Google Chrome и ChromeDriver."
            mbox.showerror("Ошибка запуска браузера", text)
            return

        driver.implicitly_wait(5)

        banner = driver.find_element(By.ID, 'bannerPopup')

        driver.execute_script("arguments[0].style.display = 'none';", banner)

        time.sleep(3)

        ruble_input_field = driver.find_element(By.CLASS_NAME, 'main-sum-field-input')

        val = debt_sum
        summa_for_work = 0

        # Массивы данных для будущей таблицы
        dates = []
        date_period = []
        num_sum = []
        summa = []
        procent_of_period = []
        day_in_month = []
        bank_rate = []
        procent_of_day = []

        currentYear = datetime.now().year
        currentDay = datetime.now().day
        currentMonth = datetime.now().month

        date_from = driver.find_element(By.ID, 'dateFrom')
        shadow_root_date_from = driver.execute_script('return arguments[0].shadowRoot', date_from)
        date_from_input = shadow_root_date_from.find_element(By.CLASS_NAME, 'dateinput')

        date_to = driver.find_element(By.ID, 'dateTo')
        shadow_root_date_to = driver.execute_script('return arguments[0].shadowRoot', date_to)
        date_to_input = shadow_root_date_to.find_element(By.CLASS_NAME, 'dateinput')

        use_date = date(first_year, first_month, first_day)

        start_date = date(first_year, first_month, first_day)

        end_date = date(first_year, first_month, first_day)

        start_period = list()
        end_period = list()

        if var_last_day:
            currentDay -= 1
            last_day = currentDay
            last_year = currentYear
            last_month = currentMonth

        for year in range(first_year, last_year + 1, 1):
            for month in range(1, 13, 1):
                if year == first_year and month < first_month:
                    continue

                if var_begin_month:
                    # блок счета с выбранного дня
                    if year == last_year and use_date.month == last_month:
                        if use_date.day == last_day:
                            break
                        if use_date.day < last_day:
                            count_day = last_day - use_date.day + 1
                            next_month_date = use_date + relativedelta(months=+1)
                            count_day_in_month = (next_month_date - use_date).days
                            use_date = use_date.replace(day=last_day)
                            val = val / count_day_in_month * count_day
                    else:
                        use_date = use_date + relativedelta(months=+1)

                    if year == last_year and use_date.month > last_month:
                        break

                    if year == last_year and use_date.month == last_month:
                        if use_date.day > last_day:
                            count_day_in_month = use_date - start_date
                            use_date = use_date.replace(day=last_day)
                            val = val / count_day_in_month.days * (use_date - start_date).days

                    end_date = use_date
                    end_date_truth = end_date + relativedelta(days=-1)

                    if year == last_year and use_date.month == last_month:
                        if use_date.day >= last_day:
                            end_date_truth = end_date + relativedelta(days=+1)
                        if var_last_day:
                            end_date_truth = end_date
                    start_period = list(reversed(start_date.isoformat().split('-')))
                    end_period = list(reversed(end_date_truth.isoformat().split('-')))
                    start_date = end_date

                else:
                    # блок счета с начала месяца

                    if year == last_year and month > last_month:
                        break

                    last_day_of_month = calendar.monthrange(use_date.year, use_date.month)[1]

                    if year == first_year and month == first_month:
                        if last_day_of_month >= first_day:
                            end_date = end_date.replace(day=last_day_of_month)
                            count_day_in_month = calendar.monthrange(end_date.year, end_date.month)[1]
                            count_days = last_day_of_month - first_day + 1
                            val = val / count_day_in_month * count_days

                    else:
                        use_date = use_date + relativedelta(months=+1)
                        val = debt_sum
                        use_date = use_date.replace(day=1)
                        start_date = start_date.replace(year=use_date.year, month=use_date.month, day=use_date.day)
                        last_day_of_month = calendar.monthrange(use_date.year, use_date.month)[1]
                        end_date = end_date.replace(year=use_date.year, month=use_date.month, day=last_day_of_month)

                    if year == last_year and use_date.month == last_month:
                        last_day_of_month = last_day
                        end_date = end_date.replace(day=last_day_of_month)
                        count_day_in_month = calendar.monthrange(end_date.year, end_date.month)[1]
                        val = val / count_day_in_month * last_day_of_month

                    start_period = list(reversed(start_date.isoformat().split('-')))
                    end_period = list(reversed(end_date.isoformat().split('-')))



                current_date_from = f"{start_period[0]}{start_period[1]}{start_period[2]}"

                date_from_input.clear()
                date_from_input.send_keys(current_date_from, Keys.ENTER)

                time.sleep(var_delay)

                current_date_to = f"{end_period[0]}{end_period[1]}{end_period[2]}"

                date_to_input.clear()
                date_to_input.send_keys(current_date_to, Keys.ENTER)

                time.sleep(var_delay)

                summa_for_work += val
                ruble_input_field.clear()
                ruble_input_field.send_keys(summa_for_work, Keys.ARROW_RIGHT)
                time.sleep(var_delay)

                try:
                    district_input = driver.find_element(By.ID, 'districtInput')
                    date_picker = district_input.find_element(By.XPATH,
                                                              '//*[@id="districtInput"]/div[2]/dropdown-select')
                    shadow_root_district = driver.execute_script('return arguments[0].shadowRoot', date_picker)
                    shadow_root_district.find_element(By.CLASS_NAME, 'button').click()
                    input_region_field = shadow_root_district.find_element(By.CLASS_NAME, 'base-input')
                    input_region_field.send_keys('Республика Бурятия', Keys.ENTER)
                except Exception as ex:
                    pass

                time.sleep(0.1)
                driver.find_element(By.ID, 'calcSubmit').click()
                time.sleep(var_delay)

                table = driver.find_element(By.CLASS_NAME, 'periods-detalization')
                inner_html_table = table.get_attribute('innerHTML')
                pandas_table = pd.read_html(StringIO(inner_html_table), thousands=None)
                pandas_table = pandas_table[0]
                for row, cell in enumerate(pandas_table.values):
                    if row > 0:
                        dates.append(format_date_with_dot(current_date_from))
                        date_period.append(cell[0])
                        num_sum.append(val)
                        summa.append(summa_for_work)
                        procent_of_period.append(float(cell[4].replace(',', '.').replace('\xa0', '')))
                        day_in_month.append(int(cell[1]))
                        bank_rate.append(float(cell[2].replace(',', '.').replace('\xa0', '')))
                        row = len(num_sum) + 1
                        procent_of_day.append(f"=F{row}/G{row}")
                time.sleep(0.1)

        dict_result = {
            'Дата месяц/год': dates,
            'период расчета': date_period,
            'начисления': num_sum,
            'недоимка': summa,
            '% за период': procent_of_period,
            'просрочка дней': day_in_month,
            'ключевая ставка': bank_rate,
            '% за день': procent_of_day
        }
        df = pd.DataFrame(dict_result)
        df.to_excel(f'{directory_for_result}/result_{debt_sum}_{datetime.now().second}.xlsx')

        time.sleep(3)
        logger.info("Successful starting of the bot")

    except:
        logger.exception("Bot error: ")
