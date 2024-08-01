import logging
import os
import smtplib
import time
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# Константы и конфигурация
GECKODRIVER_PATH = "/usr/local/bin/geckodriver"
FIREFOX_PATH = "/usr/bin/firefox"
DOWNLOAD_DIR = os.path.join(os.getcwd(), "data")
EMAIL = os.getenv("EMAIL")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
RECIPIENT_EMAIL = "gfmnlk@gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)


# Функция для вычисления первой и последней дат предыдущего месяца
def get_previous_month_dates():
    today = datetime.today()
    first_day_this_month = today.replace(day=1)
    last_day_last_month = first_day_this_month - timedelta(days=1)
    first_day_last_month = last_day_last_month.replace(day=1)
    return first_day_last_month, last_day_last_month


# Функция для очистки папки загрузки от старых файлов
def clean_download_dir(directory):
    try:
        for filename in os.listdir(directory):
            file_path = os.path.join(directory, filename)
            if os.path.isfile(file_path):
                os.unlink(file_path)
                logging.info(f"Удален файл: {file_path}")
            elif os.path.isdir(file_path):
                os.rmdir(file_path)
                logging.info(f"Удалена папка: {file_path}")
    except Exception as e:
        logging.error(f"Ошибка очистки папки {directory}\n{e}")


# Функция кликающая на элементы
def click_to_elem(element_id_tuple, err_message, timeout, driver):
    try:
        elem = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable(element_id_tuple)
        )
        elem.click()
    except Exception as e:
        logging.error(f"{err_message}\n{e}")


# Функция заполняющая поле ввода
def send_str(element_id_tuple, string, err_message, timeout, driver):
    try:
        input_field = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located(element_id_tuple)
        )
        input_field.clear()
        input_field.send_keys(string)
    except Exception as e:
        logging.error(f"{err_message}\n{e}")


# Функция для парсинга XML и получения данных
def parse_xml(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()
    data = []
    for row in root.findall(".//row[@clearing='vk']"):  # основной клиринг
        tradedate = row.get("tradedate")
        tradetime = row.get("tradetime")
        rate = float(row.get("rate"))
        data.append([tradedate, rate, tradetime])
    return data


# Получаем даты начала и конца предыдущего месяца
start_date, end_date = get_previous_month_dates()
start_date_str = start_date.strftime("%d.%m.%Y")
end_date_str = end_date.strftime("%d.%m.%Y")

# Создание папки для загрузки или очистка существующей
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
clean_download_dir(DOWNLOAD_DIR)

# Настройки Firefox
firefox_profile = webdriver.FirefoxProfile()
firefox_profile.set_preference("browser.download.folderList", 2)
firefox_profile.set_preference("browser.download.dir", DOWNLOAD_DIR)
firefox_profile.set_preference(
    "browser.helperApps.neverAsk.saveToDisk", "application/xml"
)

options = Options()
options.binary_location = FIREFOX_PATH
options.profile = firefox_profile

service = FirefoxService(executable_path=GECKODRIVER_PATH)
driver = webdriver.Firefox(service=service, options=options)
driver.get("https://www.moex.com")

time.sleep(2)

# Основной процесс парсинга
try:
    # Закрыть всплывающее окно о cookies, если оно есть
    click_to_elem(
        (
            By.XPATH,
            '//span[@class="new-ui-button__label" and text()="Принять"]',
        ),
        "Не удалось нажать на кнопку 'Принять' о куках",
        10,
        driver,
    )

    # Открыть меню
    click_to_elem(
        (By.CSS_SELECTOR, ".header__button.header-col.header-col--burger"),
        "Не удалось нажать на кнопку 'Меню'",
        10,
        driver,
    )

    time.sleep(1)

    # Переход по меню "Срочный рынок"
    click_to_elem(
        (By.LINK_TEXT, "Срочный рынок"),
        "Не удалось нажать на пункт меню 'Срочный рынок'",
        10,
        driver,
    )

    time.sleep(3)

    # Нажатие на кнопку "Согласен" с условиями использования сайта
    click_to_elem(
        (By.XPATH, '//a[@class="btn2 btn2-primary" and text()="Согласен"]'),
        "Не удалось нажать на кнопку 'Согласен'",
        10,
        driver,
    )

    # Переход по пункту "Индикативные курсы"
    click_to_elem(
        (By.LINK_TEXT, "Индикативные курсы"),
        "Не удалось нажать на пункт меню 'Индикативные курсы'",
        10,
        driver,
    )

    # Открытие выпадающего списка
    click_to_elem(
        (By.XPATH, '//div[@class="ui-select__activator -selected"]'),
        "Не удалось открыть выпадающий список валют",
        10,
        driver,
    )

    # Выбор элемента USD/RUB из списка
    click_to_elem(
        (
            By.XPATH,
            '//a[contains(text(), "USD/RUB - Доллар США к российскому рублю")]',
        ),
        "Не удалось выбрать валюту 'USD/RUB'",
        10,
        driver,
    )

    time.sleep(5)

    # Заполнение поля с начальной датой
    send_str(
        (By.ID, "fromDate"),
        start_date_str,
        "Не удалось заполнить поле начальной даты",
        10,
        driver,
    )

    # Заполнение поля с конечной датой
    send_str(
        (By.ID, "tillDate"),
        end_date_str,
        "Не удалось заполнить поле конечной даты",
        10,
        driver,
    )

    # Нажатие на кнопку "Показать"
    click_to_elem(
        (By.XPATH, '//button[@type="submit" and @aria-label="Показать"]'),
        "Не удалось нажать на кнопку 'Показать'",
        10,
        driver,
    )

    time.sleep(5)

    # Клик на ссылку для загрузки данных в XML
    click_to_elem(
        (By.XPATH, '//a[text()="Получить данные в XML"]'),
        "Не удалось нажать на ссылку для скачивания XML",
        10,
        driver,
    )

    time.sleep(10)  # Время ожидания загрузки файла

    # Переключение на другую пару валют JPY/RUB

    # Открытие выпадающего списка
    click_to_elem(
        (By.XPATH, '//div[@class="ui-select__activator -selected"]'),
        "Не удалось открыть выпадающий список валют",
        10,
        driver,
    )

    # Выбор элемента JPY/RUB из списка
    click_to_elem(
        (
            By.XPATH,
            '//a[contains(text(), "JPY/RUB - Японская йена к российскому рублю")]',
        ),
        "Не удалось выбрать валюту 'JPY/RUB'",
        10,
        driver,
    )

    time.sleep(5)

    # Заполнение поля с начальной датой
    send_str(
        (By.ID, "fromDate"),
        start_date_str,
        "Не удалось заполнить поле начальной даты",
        10,
        driver,
    )

    # Заполнение поля с конечной датой
    send_str(
        (By.ID, "tillDate"),
        end_date_str,
        "Не удалось заполнить поле конечной даты",
        10,
        driver,
    )

    # Нажатие на кнопку "Показать"
    click_to_elem(
        (By.XPATH, '//button[@type="submit" and @aria-label="Показать"]'),
        "Не удалось нажать на кнопку 'Показать'",
        10,
        driver,
    )

    time.sleep(5)

    # Клик на ссылку для загрузки данных в XML
    click_to_elem(
        (By.XPATH, '//a[text()="Получить данные в XML"]'),
        "Не удалось нажать на ссылку для скачивания XML",
        10,
        driver,
    )

    time.sleep(10)  # Время ожидания загрузки файла

finally:
    driver.quit()

# Парсинг XML и сохранение данных в Excel с помощью pandas

# Получение XML-файлов из папки
usd_rub_file = None
jpy_rub_file = None
for file in os.listdir(DOWNLOAD_DIR):
    if "USD_RUB" in file:
        usd_rub_file = os.path.join(DOWNLOAD_DIR, file)
    elif "JPY_RUB" in file:
        jpy_rub_file = os.path.join(DOWNLOAD_DIR, file)

# Парсинг данных из файлов
usd_rub_data = parse_xml(usd_rub_file) if usd_rub_file else []
jpy_rub_data = parse_xml(jpy_rub_file) if jpy_rub_file else []

# Преобразование данных в DataFrame
usd_rub_df = pd.DataFrame(
    usd_rub_data, columns=["Дата USD/RUB", "Курс USD/RUB", "Время USD/RUB"]
)
jpy_rub_df = pd.DataFrame(
    jpy_rub_data, columns=["Дата JPY/RUB", "Курс JPY/RUB", "Время JPY/RUB"]
)

# Создание итогового DataFrame и расчет столбца "Результат"
final_df = pd.concat([usd_rub_df, jpy_rub_df], axis=1)
final_df["Результат"] = (
    final_df["Курс USD/RUB"] / final_df["Курс JPY/RUB"]
).round(5)

# Запись данных в Excel
excel_file = os.path.join(DOWNLOAD_DIR, "report.xlsx")
with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
    final_df.to_excel(writer, index=False)
    workbook = writer.book
    worksheet = writer.sheets["Sheet1"]

    # Форматирование
    for column_cells in worksheet.columns:
        max_length = 0
        for cell in column_cells:
            if isinstance(cell.value, (float, int)):
                if cell.column_letter == "G":
                    cell.number_format = "[$￥-411]#,##0.00;-[$￥-411]#,##0.00"
                else:
                    cell.number_format = "#,##0.00 [$₽-419];-#,##0.00 [$₽-419]"
            cell.alignment = Alignment(horizontal="center")
            max_length = max(max_length, len(str(cell.value)))
        adjusted_width = (max_length + 2) * 1.2
        worksheet.column_dimensions[column_cells[0].column_letter].width = (
            adjusted_width
        )

# Подсчет количества строк
num_rows = len(final_df)

# Верное склонение слова "строка"
if 11 <= num_rows % 100 <= 19:
    form = "строк"
else:
    last_digit = num_rows % 10
    if last_digit == 1:
        form = "строку"
    elif 2 <= last_digit <= 4:
        form = "строки"
    else:
        form = "строк"

# Отправка письма с вложением
msg = MIMEMultipart()
msg["From"] = EMAIL
msg["To"] = RECIPIENT_EMAIL
msg["Subject"] = "Отчет с курсами валют"
body = f"Отчет содержит {num_rows} {form}."
msg.attach(MIMEText(body, "plain"))

with open(excel_file, "rb") as attachment:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {os.path.basename(excel_file)}",
    )
    msg.attach(part)

try:
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL, EMAIL_PASSWORD)
        server.sendmail(EMAIL, RECIPIENT_EMAIL, msg.as_string())
except Exception as e:
    logging.error(f"Не удалось отправить письмо\n{e}")
