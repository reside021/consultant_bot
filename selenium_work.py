import time
import calendar
import logging

from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from datetime import datetime
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService


import pandas as pd
from tkinter import messagebox as mbox

module_logger = logging.getLogger("main.selenium")


def format_date_with_dot(date):
    date = date[:2] + '.' + date[2:4] + '.' + date[4:]
    return date


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
            service = ChromeService(executable_path=ChromeDriverManager().install())
            if var_visual:
                options.add_argument("--start-maximized")
            else:
                options.add_argument('--headless')
                options.add_argument('--disable-gpu')
            driver = webdriver.Chrome(service=service, options=options)
            driver.get("https://calc.consultant.ru/395gk")
        except Exception as ex:
            text = f"Текст ошибки:\n\n{ex.msg}\n\nПроверьте версии Google Chrome и ChromeDriver."
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
        date = []
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

        for year in range(2000, currentYear + 1, 1):
            if year < first_year:
                continue
            for month in range(1, 13, 1):

                if year == first_year and month < first_month:
                    continue

                if year == currentYear and month > currentMonth:
                    break

                if year > last_year:
                    break

                if year == last_year and month > last_month:
                    break

                if year == first_year and month > first_month:
                    first_day = 1

                date_from = driver.find_element(By.ID, 'dateFrom')
                shadow_root_date_from = driver.execute_script('return arguments[0].shadowRoot', date_from)
                date_from_input = shadow_root_date_from.find_element(By.CLASS_NAME, 'dateinput')
                if first_day < 10:
                    current_date_from = f"0{first_day}"
                else:
                    current_date_from = f"{first_day}"
                if month < 10:
                    current_date_from += f"0{month}{year}"
                else:
                    current_date_from += f"{month}{year}"
                date_from_input.clear()
                date_from_input.send_keys(current_date_from, Keys.ENTER)

                time.sleep(var_delay)

                date_to = driver.find_element(By.ID, 'dateTo')
                shadow_root_date_to = driver.execute_script('return arguments[0].shadowRoot', date_to)
                date_to_input = shadow_root_date_to.find_element(By.CLASS_NAME, 'dateinput')

                last_day_of_month = calendar.monthrange(year, month)[1]

                if not var_last_day:
                    if year == last_year and month == last_month:
                        last_day_of_month = last_day
                else:
                    if year == currentYear and month == currentMonth:
                        if last_day_of_month >= currentDay:
                            last_day_of_month = currentDay - 1

                if last_day_of_month < 10:
                    current_date_to = f"0{last_day_of_month}"
                else:
                    current_date_to = f"{last_day_of_month}"
                if month < 10:
                    current_date_to += f"0{month}{year}"
                else:
                    current_date_to += f"{month}{year}"
                date_to_input.clear()
                date_to_input.send_keys(current_date_to, Keys.ENTER)
                time.sleep(var_delay)

                # if not var_last_day:
                if year == last_year and month == last_month:
                    count_day_in_month = calendar.monthrange(year, month)[1]
                    val = val / count_day_in_month * last_day_of_month

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
                pandas_table = pd.read_html(inner_html_table, thousands=None)
                pandas_table = pandas_table[0]
                for row, cell in enumerate(pandas_table.values):
                    if row >= 2:
                        date.append(format_date_with_dot(current_date_from))
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
            'Дата месяц/год': date,
            'период расчета': date_period,
            'начисления': num_sum,
            'недоимка': summa,
            '% за период': procent_of_period,
            'просрочка дней': day_in_month,
            'ключевая ставка': bank_rate,
            '% за день': procent_of_day
        }
        df = pd.DataFrame(dict_result)
        df.to_excel(f'{directory_for_result}/result_{val}.xlsx')

        time.sleep(3)
        logger.info("Successful starting of the bot")

    except:
        logger.exception("Bot error: ")
