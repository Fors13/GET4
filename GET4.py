# -*- coding: utf-8 -*-
from math import e
import tkinter as tk
from tkinter import scrolledtext, filedialog
import threading
import os
import time
import pandas as pd
import socket
import subprocess
import requests
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from zipfile import ZipFile, BadZipFile
from selenium.webdriver.support.ui import Select


# ===== Настройки Selenium =====
EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
PROFILE_PATH = r"C:\SeleniumEdgeProfile"
MAIN_URL = "https://mercury.vetrf.ru/"

# ===== Флаги управления =====
script_running = False
auth_confirmed = False

# ===== Функции =====


def log(msg, log_widget):
    log_widget.insert(tk.END, f"{msg}\n")
    log_widget.see(tk.END)


def stop_script():
    global script_running
    script_running = False


def continue_after_auth():
    global auth_confirmed
    auth_confirmed = True


def value_to_str(val):
    """Конвертация значения из Excel без .0"""
    if pd.isna(val):
        return ""
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return str(val)
    return str(val)


def format_excel_date(val):
    """Преобразует значение из Excel в строку формата DD.MM.YYYY"""
    if pd.isna(val):
        return ""
    if isinstance(val, pd.Timestamp):
        return val.strftime("%d.%m.%Y")
    if isinstance(val, str):
        try:
            dt = pd.to_datetime(val, dayfirst=True, errors='coerce')
            if pd.notna(dt):
                return dt.strftime("%d.%m.%Y")
            return val
        except:
            return val
    if isinstance(val, (float, int)):
        # Excel serial date
        try:
            dt = pd.to_datetime('1899-12-30') + pd.to_timedelta(int(val), 'D')
            return dt.strftime("%d.%m.%Y")
        except:
            return str(val)
    return str(val)


def is_port_open(port):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.connect(("127.0.0.1", port))
        s.close()
        return True
    except:
        return False


def check_site_available(url):
    try:
        response = requests.get(url, timeout=5)
        return response.status_code == 200
    except:
        return False


def send_email_notification(subject, body, log_widget=None, to_emails=["Evgeny.Bardin@hochland.ru"], from_email="Evgeny.Bardin@hochland.ru"):
    try:
        smtp_server = "hl-smtp.hochland.com"
        smtp_port = 25
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = ", ".join(to_emails)
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.sendmail(from_email, to_emails, msg.as_string())

        if log_widget:
            log(f"[EMAIL] Сообщение успешно отправлено: {subject}", log_widget)
    except Exception as e:
        if log_widget:
            log(f"[EMAIL] Ошибка при отправке письма: {e}", log_widget)


def is_valid_xlsx(file_path):
    try:
        with ZipFile(file_path, 'r') as zip_file:
            return True
    except BadZipFile:
        return False


def get_column_by_keywords(df, keywords):
    for name in df.columns:
        for kw in keywords:
            if kw.lower() in name.lower():
                return name
    return None


DEBUG_PORT = 9222  # пример порта
EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
PROFILE_PATH = r"C:\EdgeProfile"


def start_edge_if_needed(log_widget):
    """Запускает Microsoft Edge с удалённой отладкой, если не запущен, и возвращает драйвер Selenium."""

    edge_options = Options()
    edge_options.debugger_address = f"127.0.0.1:{DEBUG_PORT}"

    if not is_port_open(DEBUG_PORT):
        log("Edge не запущен. Запускаем новый браузер...", log_widget)
        try:
            subprocess.Popen(
                [EDGE_PATH, f"--remote-debugging-port={DEBUG_PORT}", f"--user-data-dir={PROFILE_PATH}"])
            log("Браузер Edge запускается...", log_widget)
            time.sleep(5)  # ждём запуска
        except Exception as e:
            log(f"Ошибка при запуске Edge: {e}", log_widget)
            return None
    else:
        log("Edge уже запущен. Подключаемся к существующему сеансу...", log_widget)

    try:
        driver = webdriver.Edge(options=edge_options)
        log("Успешно подключились к Edge.", log_widget)
        return driver
    except Exception as e:
        log(f"Не удалось подключиться к Edge: {e}", log_widget)
        return None


def go_to_all_records(driver, log_widget):
    try:
        if not check_site_available(MAIN_URL):
            log("Сайт недоступен! Ждём 5 секунд...", log_widget)
            time.sleep(5)
        journal_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//a[contains(@href, "operatorui?_action=listRealTrafficVU")]'))
        )
        journal_link.click()
        time.sleep(2)
        all_records_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//a[@href=\'javascript:onEnterMenu("", "AjaxAllRealTrafficVU", "");\']'))
        )
        all_records_link.click()
        time.sleep(2)
    except Exception as e:
        log(
            f"Ошибка при переходе к 'Журнал продукции' / 'Все записи': {e}", log_widget)
        send_email_notification("Ошибка скрипта", str(e), log_widget)


# Проверка доступности кнопки "Транзакции" нажатие на нее
def go_to_all_records2(driver, log_widget):
    try:
        if not check_site_available(MAIN_URL):
            log("Сайт недоступен! Ждём 5 секунд...", log_widget)
            time.sleep(5)
        transactions_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, 'operatorui?_action=listTransaction')]")))
        transactions_link.click()
        time.sleep(2)
        all_records_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(@href, 'onEnterMenu') and contains(text(), 'Шаблоны')]"))
        )
        all_records_link.click()
        time.sleep(2)
    except Exception as e:
        log(
            f"Ошибка при переходе к 'Транзакции' / 'Шаблоны': {e}", log_widget)
        send_email_notification("Ошибка скрипта", str(e), log_widget)


def lab_results_table_exists(driver, log_widget=None):
    """Проверка: таблица исследований существует и содержит строки"""
    try:
        table = driver.find_element(
            By.XPATH, '//h4[contains(text(),"Лабораторные исследования")]/following-sibling::table[contains(@class,"innerForm")]')
        rows = table.find_elements(By.TAG_NAME, "tr")
        if len(rows) > 1:
            return True
        return False
    except NoSuchElementException:
        return False

# ===== Основная функция Заполнение Лаб Данных  =====


def main(export_file, lab_file, log_widget):
    global script_running, auth_confirmed
    script_running = True
    auth_confirmed = False
    errors = []

    try:
        if not export_file or not os.path.exists(export_file):
            log("Файл EXPORT не выбран или отсутствует!", log_widget)
            return
        if not lab_file or not os.path.exists(lab_file):
            log("Файл 'Экспорт лаб ис' не выбран или отсутствует!", log_widget)
            return
        if not is_valid_xlsx(export_file):
            log("Файл EXPORT повреждён или не является XLSX!", log_widget)
            return
        if not is_valid_xlsx(lab_file):
            log("Файл 'Экспорт лаб ис' повреждён или не является XLSX!", log_widget)
            return

        df_export = pd.read_excel(export_file, engine="openpyxl")
        df_lab = pd.read_excel(lab_file, engine="openpyxl")

        col_num = get_column_by_keywords(
            df_export, ["Номер записи склад.журнала"])
        col_material = get_column_by_keywords(df_export, ["Материал"])
        col_lab_name = get_column_by_keywords(
            df_lab, ["Наименование лаборатории"])
        col_disease = get_column_by_keywords(
            df_lab, ["Наименование показателя"])
        col_research_date = get_column_by_keywords(
            df_lab, ["Дата получения результата"])
        col_expertise = get_column_by_keywords(df_lab, ["№ экспертизы"])
        col_lab_material = get_column_by_keywords(df_lab, ["Артикул"])

        driver = start_edge_if_needed(log_widget)
        driver.get(MAIN_URL)
        log("Авторизуйся в браузере и нажми кнопку 'Я авторизовался — продолжить' в GUI...", log_widget)

        while not auth_confirmed and script_running:
            time.sleep(1)
        if not script_running:
            log("Скрипт остановлен до начала обработки.", log_widget)
            return
# =====Запуск цыкла по заполнению лаб.данных=====
        for idx, row in df_export.iterrows():
            if not script_running:
                log("Цикл остановлен пользователем.", log_widget)
                break
            try:
                num_sj = value_to_str(row[col_num]) if col_num else ""
        # Если номер записи не указан — пропускаем
                if not num_sj or str(num_sj).strip() == "":
                    log(f"Пропуск строки {idx + 1}: поле 'Номер записи склад.журнала' не заполнено.", log_widget)
                    continue
        ########################
                material = value_to_str(
                    row[col_material]) if col_material else ""
                log(f"\nОбрабатываем запись: {num_sj}, материал: {material}", log_widget)

                go_to_all_records(driver, log_widget)

# нажатие кнопки "поиск"
                search_button = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.ID, "findFormTop"))
                )
                driver.execute_script("arguments[0].click();", search_button)
                time.sleep(1)

# Заполнение поля ЗСЖ в поиске
                search_box = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.ID, "findFormId"))
                )
                search_box.clear()
                search_box.send_keys(num_sj)
                search_box.send_keys(Keys.RETURN)
                time.sleep(2)

# Кнопка "Просмотр" найденной запаси
                view_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, 'img[alt="Просмотр"]'))
                )
                driver.execute_script("arguments[0].click();", view_button)
                time.sleep(2)

# Проверка существующей записи (наличие таблицы)
                if lab_results_table_exists(driver, log_widget):
                    log(f"У записи {num_sj} уже есть таблица исследований. Пропускаем...", log_widget)
                    continue

# Нажатие кнопки "Добавить"
                add_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, '//h4[contains(.,"Лабораторные исследования")]/a'))
                )
                driver.execute_script("arguments[0].click();", add_button)
                time.sleep(2)
# Очищаем поле "Дата акта"
                act_date_field = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.ID, "actDate"))
                )
                act_date_field.clear()

# Проверка артикула в файле, если не найден, пропускаем запись
                lab_rows = df_lab[df_lab[col_lab_material].astype(
                    str) == material]
                if lab_rows.empty:
                    log(f" Для материала {material} не найдены данные в 'Экспорт лаб ис'. Пропуск...", log_widget)
                    continue
# Заполняем поля
                lab_row = lab_rows.iloc[0]

                lab_name = value_to_str(
                    lab_row[col_lab_name]) if col_lab_name else ""
                disease = value_to_str(
                    lab_row[col_disease]) if col_disease else ""
                research_date = format_excel_date(
                    lab_row[col_research_date]) if col_research_date else ""
                expertise_number = value_to_str(
                    lab_row[col_expertise]) if col_expertise else ""

                driver.find_element(By.NAME, "laboratory").send_keys(lab_name)
                driver.find_element(By.NAME, "disease").send_keys(disease)
                research_field = driver.find_element(By.ID, "researchDate")
                research_field.clear()
                research_field.send_keys(research_date)
                driver.find_element(By.NAME, "expertiseNumber").send_keys(
                    expertise_number)

                result_select = driver.find_element(By.NAME, "result")
                for option in result_select.find_elements(By.TAG_NAME, "option"):
                    if "отрицательный" in option.text.lower():
                        option.click()
                        break

                driver.find_element(
                    By.NAME, "conclusion").send_keys("Отрицательно")

                save_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, 'button.positive'))
                )

# Сохронение записи
                save_button.click()
                time.sleep(2)

                log(f"Запись {num_sj} успешно обработана.", log_widget)

            except Exception as e:
                log(f"Ошибка при обработке записи {num_sj}: {e}", log_widget)
                send_email_notification(
                    subject="Ошибка обработки записи",
                    body=f"Ошибка при обработке записи {num_sj}: {e}",
                    log_widget=log_widget
                )
                errors.append(num_sj)
                continue
        log("Скрипт завершён.", log_widget)

    # except Exception as e:
        # log(f"Фатальная ошибка: {e}", log_widget)
    # Сбор ошибок и вывод в агрегированном состоянии
    finally:

        if errors:
            log("\n=== Ошибки при обработке следующих записей ===", log_widget)
            for num in errors:
                log(f" - {num}", log_widget)
        else:
            log("Ошибок не обнаружено — все записи обработаны успешно.", log_widget)

        # закрываем драйвер аккуратно, если он был создан
        try:
            if driver:
                driver.quit()
        except Exception:
            pass

            # ===== Основная функция Создание ВСД  =====


def main2(export_file, log_widget, skip_date):

    global script_running, auth_confirmed
    script_running = True
    auth_confirmed = False
    errors = []

# Проверка файла
    if not export_file or not os.path.exists(export_file):
        log("Файл EXPORT не выбран или отсутствует!", log_widget)
        return

    if not is_valid_xlsx(export_file):
        log("Файл EXPORT повреждён или не является XLSX!", log_widget)
        return

# Чтение Excel
    df_export = pd.read_excel(export_file, engine="openpyxl", dtype=str)

# Берём только первую строку
    row = df_export.iloc[0]
# Записи из файлов
    col_nam_pet = get_column_by_keywords(df_export, ["Название шаблона"])
    col_mat = get_column_by_keywords(df_export, ["Материал"])
    col_nam_matl = get_column_by_keywords(df_export, ["Название материала"])
    col_order = get_column_by_keywords(df_export, ["Документ-образец"])
    col_num = get_column_by_keywords(df_export, ["Номер записи склад.журнала"])
    col_nweight = get_column_by_keywords(df_export, ["Вес нетто"])
    col_volumе_supply = get_column_by_keywords(df_export, ["Объем поставки"])
    col_party = get_column_by_keywords(df_export, ["Партия"])
    col_special_notes = get_column_by_keywords(df_export, ["Особые отметки"])
    col_series_secure_form = get_column_by_keywords(
        df_export, ["Cерия защищенного бланка"])
    col_secure_form_number = get_column_by_keywords(
        df_export, ["Номер защищенного бланка"])
    col_numauto = get_column_by_keywords(df_export, ["Номер машины"])
    col_trailer = get_column_by_keywords(df_export, ["Номер прицепа"])

    nam_pet = value_to_str(row[col_nam_pet]) if col_nam_pet else ""
    mat = value_to_str(row[col_mat]) if col_mat else ""
    nam_matl = value_to_str(row[col_nam_matl]) if col_nam_matl else ""
    order = value_to_str(row[col_order]) if col_order else ""
    numauto = value_to_str(row[col_numauto]) if col_numauto else ""
    num_trailer = value_to_str(row[col_trailer]) if col_trailer else ""
    cmr = value_to_str(row[col_order]) if col_order else ""
    ttn = value_to_str(row[col_order]) if col_order else ""
    log(f"Обрабатываем запись: {nam_pet}, Номер машины: {numauto}, Номер прицепа: {num_trailer}", log_widget)

# Запуск браузера
    driver = start_edge_if_needed(log_widget)
    driver.get(MAIN_URL)
    log("Авторизуйся в браузере и нажми кнопку 'Я авторизовался — продолжить' в GUI...", log_widget)

# Ждём авторизации
    while not auth_confirmed and script_running:
        time.sleep(1)

    if not script_running:
        log("Скрипт остановлен до начала обработки.", log_widget)
        return

    try:
        # Переход в раздел "Транзакции"
        go_to_all_records2(driver, log_widget)

# Открыть поиск
        search_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "findFormTop"))
        )
        driver.execute_script("arguments[0].click();", search_button)
        time.sleep(1)

# Заполнить поле "Название шаблона"
        search_box = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, "templateName"))
        )
        search_box.clear()
        search_box.send_keys(nam_pet)
        time.sleep(2)
        search_box.send_keys(Keys.RETURN)
        time.sleep(2)

# Нажатие кнопки "Просмотр"
        view_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, 'img[alt="Просмотр"]'))
        )
        driver.execute_script("arguments[0].click();", view_button)
        time.sleep(2)

# Нажатие кнопки "Редактировать"
        edit_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(@onclick, 'modifyTransactionForm')]"))
        )
        driver.execute_script("arguments[0].click();", edit_button)
        time.sleep(2)

# Заполнение "Номер авто"
        numauto_box = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.NAME, "transportAuto"))
        )
        numauto_box.clear()
        numauto_box.send_keys(numauto)
        numauto_box.send_keys(Keys.RETURN)

# Заполнение "Номер прицепа"
        trailero_box = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.NAME, "trailer"))
        )
        trailero_box.clear()
        trailero_box.send_keys(num_trailer)
        trailero_box.send_keys(Keys.RETURN)
        time.sleep(2)

# Сохранение
        save_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.positive'))
        )
        save_button.click()
        time.sleep(2)
    except Exception:
        log(f" Не удалось подтвердить сохранение записи {nam_pet}", log_widget)

# Проверка успешного сохранения
#    try:
#        success_message = WebDriverWait(driver, 10).until(
#                EC.visibility_of_element_located((By.CSS_SELECTOR, "div.message.message-success"))
#        )
#        message_text = success_message.text.strip()
#        if "Информация успешно сохранена" in message_text:
#            log(f" Шаблон {nam_pet} успешно сохранён!", log_widget)
#        else:
#            log(f" Сообщение появилось, но текст другой: '{message_text}'", log_widget)
#    except Exception:
#            log(f" Не удалось подтвердить сохранение записи, требуется проверить шаблон {nam_pet}", log_widget)

# Нажатие кнопки "редактировать сведения"
    edit_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH,
             "//a[contains(@onclick, 'modifyWaybillForm') and @title='редактировать сведения']")
        )
    )
# Заполнение данными
    driver.execute_script("arguments[0].click();", edit_button)

    ttn_box = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.NAME, "waybillNumber"))
    )
    ttn_box.clear()
    ttn_box.send_keys(ttn)
    ttn_box.send_keys(Keys.RETURN)

    date = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[@href=\"javascript:setToday('waybillDate');\"]")))
    date.click()

# Заполнение CMR
    cmr_box = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located(
            (By.NAME, "relatedDocumentIssueNumber_1"))
    )
    cmr_box.clear()
    cmr_box.send_keys(f"CMR {cmr}")
    cmr_box.send_keys(Keys.RETURN)

    date1 = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[@href=\"javascript:setToday('issueDate_1');\"]")))
    date1.click()

# Сохраняем
    wait = WebDriverWait(driver, 10)
    save_button = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[contains(@class, 'positive') and contains(., 'Сохранить')]")))
    save_button.click()

# Проверка успешного сохранения
#    try:
#        success_message = WebDriverWait(driver, 10).until(
#            EC.visibility_of_element_located((By.CSS_SELECTOR, "div.message.message-success"))
#        )
#        message_text = success_message.text.strip()
#        if "Информация успешно сохранена" in message_text:
#            log(f" Сведения шаблона {nam_pet} успешно сохранены!", log_widget)
#        else:
#            log(f" Не удалось подтвердить сохранение записи, требуется проверить сведения шаблона '{message_text}'", log_widget)
#    except Exception:
#            log(f" Не удалось подтвердить сохранение записи, требуется проверить сведения шаблона  {nam_pet}", log_widget)

# Создать транзакцию
    wait = WebDriverWait(driver, 10)
    Create_a_transaction = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[contains(@onclick, 'generateTransactionFromTemplate')]")))
    Create_a_transaction.click()
    time.sleep(2)

# Добавить продукцию из журнала для создания ВСД
    Add_products_from_VSD = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(@onclick, 'addVetDocumentForm')]"))
    )
    # Прокручиваем к элементу (часто помогает, если страница длинная)
    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center'});", Add_products_from_VSD)
    # Кликаем
    Add_products_from_VSD.click()
    time.sleep(2)

# Нажатие "не скоропортящаяся продукция"
    non_perishable_products = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//input[@name='perishable' and @value='withoutAnimal']"))
    )
    driver.execute_script("arguments[0].click();", non_perishable_products)

# =====Цикл по добавлению из журнала для оформления ВСД=====
    for idx, row in df_export.iterrows():
        if not script_running:
            log("Цикл остановлен пользователем.", log_widget)
            break

        try:
            num_sj = value_to_str(row[col_num]) if col_num else ""
# Если номер записи не указан — пропускаем
            if not num_sj or str(num_sj).strip() == "":
                log(f"Пропуск строки {idx + 1}: поле 'Номер записи склад.журнала' не заполнено.", log_widget)
                continue
    ######################
            nweight = value_to_str(row[col_nweight]) if col_nweight else ""
            volumе_supply = value_to_str(
                row[col_volumе_supply]) if col_volumе_supply else ""
            party = value_to_str(row[col_party]) if col_party else ""
            special_notes = value_to_str(
                row[col_special_notes]) if col_special_notes else ""
            series_secure_form = value_to_str(
                row[col_series_secure_form]) if col_series_secure_form else ""
            secure_form_number = value_to_str(
                row[col_secure_form_number]) if col_secure_form_number else ""

# Заполнение ЗСЖ для поиска
            journal_entry = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located(
                    (By.NAME, "realTrafficVUTemplate"))
            )
            journal_entry.clear()
            journal_entry.send_keys(num_sj)
            journal_entry.send_keys(Keys.RETURN)
            time.sleep(2)
# Проваливаемся в строку
            # click_box = WebDriverWait(driver, 10).until(
            # EC.element_to_be_clickable((By.XPATH, "//tr[contains(@onclick, 'chooseTraffic(')]"))
            # )
            # click_box.click()
            # time.sleep(4)

# Получаем значение из страницы для проверки партии

            MAX_ATTEMPTS = 3  # максимальное число попыток

            for attempt in range(MAX_ATTEMPTS):
                # Проваливаемся в строку
                click_box = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//tr[contains(@onclick, 'chooseTraffic(')]"))
                )
                driver.execute_script("arguments[0].click();", click_box)
                time.sleep(1)  # небольшая пауза, чтобы страница обновилась

                # Получаем значение партии
                try:
                    party_td = WebDriverWait(driver, 10).until(
                        lambda d: d.find_element(
                            By.XPATH,
                            "//td[contains(text(),'Номер производственной партии')]/following-sibling::td[contains(@class,'value')]"
                        )
                    )
                    party_value = party_td.text.strip()
                except:
                    party_value = ""

                log(f"Попытка {attempt+1}: партия на странице = '{party_value}', ожидаем = '{party}'", log_widget)

                # Сравниваем с ожидаемой партией
                if party_value == str(party).strip():
                    log(f"Партия совпала: {party_value}", log_widget)
                    break  # выходим из цикла и продолжаем обработку
                else:
                    log("Партия не совпадает, повторяем клик...", log_widget)
                    time.sleep(1)
            else:
                log(f"❌ После {MAX_ATTEMPTS} попыток партия не совпала для записи {num_sj}, пропускаем.", log_widget)
                errors.append(num_sj)
                continue  # пропускаем эту запись


# Заполнить обьем нетто
            nweight_box = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.NAME, "inputWeight"))
            )
            nweight_box.clear()
            nweight_box.send_keys(nweight)
            nweight_box.send_keys(Keys.RETURN)

# Удаляем старую упаковку
            while True:
                try:
                    # Ищем кнопку удаления
                    delete_button = driver.find_element(
                        By.CSS_SELECTOR, "a.icon-link.remove i.fa-trash")

                    # Кликаем по ней (через JS, чтобы избежать "element not clickable")
                    driver.execute_script(
                        "arguments[0].click();", delete_button)
                    # Небольшая пауза, чтобы страница успела обновиться
                    time.sleep(1)

                except NoSuchElementException:
                    # Кнопки больше нет — выходим из цикла
                    print("Все кнопки удалены.")
                    break

# Кнопка "Добавить упаковку"
            add_packaging = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//a[@href='javascript:openAddBlock();' and text()='Добавить']"))
            )
            add_packaging.click()

# Клик по видимому полю уровень
            level = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.ID, "select2-levelList-container"))
            )
            level.click()
    # Выбираем нужный пункт по названию
            option = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//li[contains(@class,'select2-results__option') and text()='Транспортный (Логистический) уровень']"))
            )
            option.click()
# Клик по видимому полю упаковка
            packing = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.ID, "select2-packingList-container"))
            )
            packing.click()

    # Выбираем нужный пункт по value
            option1 = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//li[contains(@class,'select2-results__option') and text()='Коробка (BX)']"))
            )
            option1.click()
# Заполняем "Кол-во единиц упаковки"
            volumе_supply_box = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.NAME, "packingQuantity"))
            )
            volumе_supply_box.clear()
            volumе_supply_box.send_keys(volumе_supply)
            volumе_supply_box.send_keys(Keys.RETURN)

    # Кнопка "Добавить"
            add_link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//a[@href='javascript:addMarking();' and text()='Добавить']"))
            )
            add_link.click()
            time.sleep(2)

# Поле "Наименование"Очищаем и вводим текст "Партия"
            marking_input = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.ID, "markingName_0"))
            )
            marking_input.clear()
            marking_input.send_keys(party)

# Клик по видимому полю "Тип"
            select_box = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.ID, "select2-markingTypeList_0-container"))
            )
            select_box.click()
            time.sleep(0.5)
            option = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//li[contains(@class,'select2-results__option') and text()='Номер партии (BN)']"))
            )
            option.click()

# Сохранить
            save_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "savePacking"))
            )
            save_button.click()
            log(f" Упаковка записи {num_sj} успешно сохранена!", log_widget)

# # Заполняем цель
            select_el = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "target"))
            )
            Select(select_el).select_by_value(
                "639")   # реализация в пищу людям
# Изготовлена из сырья, прошедшего ветеринарно-санитарную
            radio = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "input[name='vetExpertise'][value='2']"))
            )
            radio.click()
# Особые отметки, заполнение
            special_notes_box = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.NAME, "specialMarks"))
            )
            special_notes_box.clear()
            special_notes_box.send_keys(special_notes)

# # Cерия защищенного бланка
            series_secure_form_box = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located(
                    (By.NAME, "vetDocumentSeries"))
            )
            series_secure_form_box.clear()
            series_secure_form_box.send_keys(series_secure_form)

# # Номер защищенного бланка
            secure_form_number_box = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located(
                    (By.NAME, "vetDocumentNumber"))
            )
            secure_form_number_box.clear()
            secure_form_number_box.send_keys(secure_form_number)

# Дата защищенного бланка (если не отключена галочкой)
            if not skip_date:
                date = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//a[@href=\"javascript:setToday('vetDocumentDate');\"]")))
                date.click()
            else:
                log("Пропуск заполнения даты защищенного бланка (по настройке).", log_widget)

#  ===   Ожидаем нажатие "Я авторизовался — продолжить"   ===
            # log("Перед сохранением нажмите 'Я авторизовался — продолжить'", log_widget)
            # auth_confirmed = False
            # while not auth_confirmed and script_running:
            #     time.sleep(1)
            # if not script_running:
            #     log("Скрипт остановлен пользователем перед сохранением.", log_widget)
            #     return


# Сохранение
            save_box = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(@class, 'positive') and contains(., 'Сохранить и добавить еще')]"))
            )
            save_box.click()
            time.sleep(1)


# Проверка успешного сохранения
#            try:
#                success_message = WebDriverWait(driver, 10).until(
#                    EC.visibility_of_element_located((By.CSS_SELECTOR, "div.message.message-success"))
#                )
#                message_text = success_message.text.strip()
#                if "Информация успешно сохранена" in message_text:
#                    log(f" Сведения шаблона {num_sj} успешно сохранены!", log_widget)
#                else:
#                    log(f" Не удалось подтвердить сохранение записи, требуется проверить сведения шаблона '{message_text}'", log_widget)
#            except Exception:
#                    log(f" Не удалось подтвердить сохранение записи, требуется проверить сведения шаблона  {num_sj}", log_widget)

            log(f"Запись {num_sj} успешно обработана.", log_widget)
        except Exception as e:
            log(f"Ошибка при обработке записи {num_sj}: {e}", log_widget)
            send_email_notification(
                subject="Ошибка обработки записи",
                body=f"Ошибка при обработке записи {num_sj}: {e}",
                log_widget=log_widget
            )
            errors.append(num_sj)
            continue
    log("Скрипт завершён.", log_widget)

    if errors:
        log("\n=== Ошибки при обработке следующих записей ===", log_widget)
        for num in errors:
            log(f" - {num}", log_widget)
    else:
        log("Ошибок не обнаружено — все записи обработаны успешно.", log_widget)

    # закрываем драйвер аккуратно, если он был создан
    try:
        if driver:
            driver.quit()
    except Exception:
        pass

        # log(f"Фатальная ошибка при обработке записи: {e}", log_widget)

        ###################################################################


# ===== GUI =====


def run_script_thread(export_file, lab_file, log_widget):
    threading.Thread(target=main, args=(
        export_file, lab_file, log_widget), daemon=True).start()


def run_script_thread2(export_file, log_widget):
    threading.Thread(target=main2, args=(
        export_file, log_widget, skip_date_var.get()), daemon=True).start()


def select_export_file():
    filename = filedialog.askopenfilename(title="Выберите EXPORT файл", filetypes=[
                                          ("Excel files", "*.xlsx *.xls")])
    export_file_var.set(filename)


def select_lab_file():
    filename = filedialog.askopenfilename(
        title="Выберите 'Экспорт лаб ис'", filetypes=[("Excel files", "*.xlsx *.xls")])
    lab_file_var.set(filename)


root = tk.Tk()
root.title("Mercury Automation GUI")

export_file_var = tk.StringVar()
lab_file_var = tk.StringVar()
skip_date_var = tk.BooleanVar(value=False)

tk.Label(root, text="EXPORT файл:").pack()
tk.Entry(root, textvariable=export_file_var, width=80).pack()
tk.Button(root, text="Выбрать файл", command=select_export_file).pack(pady=5)

tk.Label(root, text="'Экспорт лаб ис':").pack()
tk.Entry(root, textvariable=lab_file_var, width=80).pack()
tk.Button(root, text="Выбрать файл", command=select_lab_file).pack(pady=5)

tk.Button(root, text="Запустить скрипт", command=lambda: run_script_thread(
    export_file_var.get(), lab_file_var.get(), log_text)).pack(pady=10)
tk.Button(root, text="Создать ВСД", command=lambda: run_script_thread2(
    export_file_var.get(), log_text)).pack(pady=10)
tk.Button(root, text="Остановить цикл", command=stop_script).pack(pady=5)
tk.Button(root, text="Я авторизовался — продолжить",
          command=continue_after_auth).pack(pady=5)
tk.Checkbutton(root, text="Не заполнять дату защищенного бланка",
               variable=skip_date_var).pack(pady=5)

log_text = scrolledtext.ScrolledText(root, width=100, height=25)
log_text.pack(padx=10, pady=10)

root.mainloop()
