import os
import logging
from logging.handlers import RotatingFileHandler

import pandas as pd
from time import sleep
import re
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service
import random
from fake_useragent import FakeUserAgent
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

KEYS_LIST = ("Тип", "Площадь", "Количество комнат", "№ этажа", "Количество этажей",
             "Этажность", "Тип дома", "Материал стен", "Категория",
             "Разрешенное использование", "Кадастровый (условный) номер", "Адрес",
             "Год постройки", "Коммуникации и их характеристики",
             "Ограничения и обременения")

log_formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(funcName)s - %(message)s')
my_handler = RotatingFileHandler("parser.log", mode='a', maxBytes=2 * 1024 * 1024,
                                 backupCount=1, encoding="utf8", delay=0)
my_handler.setFormatter(log_formatter)
my_handler.setLevel(logging.INFO)

stream_handler = logging.StreamHandler()
stream_handler.setFormatter(log_formatter)
stream_handler.setLevel(logging.DEBUG)


app_log = logging.getLogger(__name__)
app_log.setLevel(logging.DEBUG)
app_log.addHandler(my_handler)
app_log.addHandler(stream_handler)


def click_button(xpath: str, timeout=3) -> None:
    """
    Функция для нажатия на кнопку
    :param timeout: время ожидания
    :param xpath: путь к кнопке
    :return: None
    """
    cur_button = WebDriverWait(browser, timeout).until(
        EC.presence_of_element_located((By.XPATH, xpath))
    )
    cur_button.click()


def set_viewport_size(driver, width, height):
    window_size = driver.execute_script("""
        return [window.outerWidth - window.innerWidth + arguments[0],
          window.outerHeight - window.innerHeight + arguments[1]];
        """, width, height)
    driver.set_window_size(*window_size)


def random_mouse_movements(driver):
    app_log.debug("Передвигаю курсор на рандомные точки...")
    for _ in range(30):
        try:
            x = random.randint(1 * _, 10 * _)
            y = random.randint(2 * _, 11 * _)
            ActionChains(driver) \
                .move_by_offset(x, y) \
                .perform()
        except Exception:
            continue


if __name__ == '__main__':
    # url_page = input("Введите ссылку для парсинга: ")
    with open("data.txt") as file:
        URLS = list()
        for line in file.readlines():
            URLS.append(line.strip())
    app_log.info("Данные из data.txt успешно загружены!")

    options = webdriver.ChromeOptions()
    options.add_argument("--disable-infobars")
    options.add_argument("--headless=new")
    options.add_argument("--incognito")
    options.add_argument("start-maximized")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument('log-level=3')
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    user_agent = FakeUserAgent().chrome
    options.add_argument(f'--user-agent={user_agent}')
    browser = webdriver.Chrome(service=Service(executable_path="./chromedriver.exe"),
                               options=options)
    for url_page in URLS:
        try:
            app_log.debug(f"Беру в обработку {url_page}...")
            num_page = url_page.split('/')[-1]
            app_log.info(f"Номер текущего лота: {num_page}")

            app_log.info("Начинаю парсинг сайта...")

            try:
                browser.get(url_page)
            except Exception:
                browser.get(url_page)

            try:
                click_button('/html/body/app-root/app-banner-stack/div/div/button')
            except Exception:
                pass  # Нет кнопки принять куки

            # Данные из карточки лота на сайте www.torgiasv.ru
            app_log.info("Приступаю к парсингу данных на www.torgiasv.ru")
            sleep(3)
            div_tag_number = 2  # В xpath путях номер div элемента меняется временами
            try:
                lot_name = WebDriverWait(browser, 1).until(
                    EC.presence_of_element_located((By.XPATH,
                                                    '/html/body/app-root/div/main/app-page-sales-item/app-catalog'
                                                    '-detail/'
                                                    'div/div[2]'
                                                    '/div/div[1]/div/app-block-detail-title/div/div[3]'))).text
            except Exception:
                try:
                    lot_name = WebDriverWait(browser, 1).until(
                        EC.presence_of_element_located((By.XPATH,
                                                        '/html/body/app-root/div/main/app-page-sales-item/app-catalog'
                                                        '-detail/'
                                                        f'div/div[{div_tag_number}]/'
                                                        'div/div[1]/div/app-block-detail-title/div/div[1]/'
                                                        'div[1]/h1'))).text
                except Exception:
                    div_tag_number += 1
                    lot_name = WebDriverWait(browser, 1).until(
                        EC.presence_of_element_located((By.XPATH,
                                                        '/html/body/app-root/div/main/app-page-sales-item/app-catalog'
                                                        '-detail/'
                                                        f'div/div[{div_tag_number}]/div/div[1]/div/'
                                                        f'app-block-detail-title/div/div[3]'))).text

            try:
                lot_number = WebDriverWait(browser, 1).until(
                    EC.presence_of_element_located(
                        (By.XPATH, '/html/body/app-root/div/main/app-page-sales-item/app-catalog-detail/'
                                   'div/div[2]/div/div[1]/div/app-block-detail-title/'
                                   'div/div[2]/div/span[2]'))).text
            except Exception:
                lot_number = WebDriverWait(browser, 1).until(
                    EC.presence_of_element_located(
                        (By.XPATH,
                         f'/html/body/app-root/div/main/app-page-sales-item/app-catalog-detail/'
                         f'div/div[{div_tag_number}]/div/div[1]/div/app-block-detail-title/'
                         f'div/div[2]/a/span[2]'))).text

            lot_number = int(lot_number)

            price = lot_name
            if "\u2009" in price:
                price = "".join(price.split("\u2009"))
                price = "".join(price.split())
                price: str = re.findall(r'(?<![\d.])\d{1,3}(?:\d{3})*,\d+', price)[0]
            else:
                price: str = re.findall(r'(?<![\d.])\d{1,3}(?:\ \d{3})*,\d+', price)[0]
            price: float = float("".join(price.split()).replace(',', '.'))

            type_of_credit = WebDriverWait(browser, 1).until(
                EC.presence_of_element_located((By.XPATH,
                                                '/html/body/app-root/div/main/app-page-sales-item/'
                                                f'app-catalog-detail/div/div[{div_tag_number}]'
                                                '/div/div[1]/div/div[3]/app-block-feature-group/div/div[2]/div/'
                                                'app-block-feature-list/div[1]/div[2]/app-block-feature-value/'
                                                'span'))).text

            type_of_property = WebDriverWait(browser, 1).until(
                EC.presence_of_element_located((By.XPATH,
                                                '/html/body/app-root/div/main/app-page-sales-item/'
                                                'app-catalog-detail/div/div[1]/div[1]/a'))).text

            region = WebDriverWait(browser, 1).until(
                EC.presence_of_element_located((By.XPATH,
                                                '/html/body/app-root/div/main/app-page-sales-item/'
                                                f'app-catalog-detail/div/div[{div_tag_number}]/div/div[1]/div/div[4]/'
                                                'app-block-feature-group/div[1]/div[2]/div/'
                                                'app-block-feature-list/div/div[2]/app-block-feature-value/span'))).text

            agency = WebDriverWait(browser, 1).until(
                EC.presence_of_element_located((By.XPATH,
                                                '/html/body/app-root/div/main/app-page-sales-item/'
                                                f'app-catalog-detail/div/div[{div_tag_number}]/div/div[1]/div/div[4]/'
                                                'app-block-feature-group/div[2]/div[2]/div/app-block-feature-list/'
                                                'div[1]/'
                                                'div[2]/app-block-feature-value/a'))).text

            data_published = WebDriverWait(browser, 1).until(
                EC.presence_of_element_located((By.XPATH,
                                                '/html/body/app-root/div/main/app-page-sales-item/app-catalog-detail/'
                                                f'div/div[{div_tag_number}]/div/div[1]/'
                                                f'div/div[4]/app-block-feature-group/div[2]/div[2]/'
                                                'div/app-block-feature-list/div[2]/div[2]/'
                                                'app-block-feature-value/span'))).text

            public_link = WebDriverWait(browser, 1).until(
                EC.presence_of_element_located((By.XPATH,
                                                '/html/body/app-root/div/main/app-page-sales-item/app-catalog-detail/'
                                                f'div/div[{div_tag_number}]/div/'
                                                f'div[1]/div/div[4]/app-block-feature-group/div[2]/div[2]/'
                                                'div/app-block-feature-list/div[3]/div[2]/'
                                                'app-block-feature-value/a'))).get_attribute('href')
            try:
                try:
                    platform = WebDriverWait(browser, 1).until(
                        EC.presence_of_element_located((By.XPATH,
                                                        '/html/body/app-root/div/main/app-page-sales-item/'
                                                        'app-catalog-detail/div/'
                                                        f'div[{div_tag_number}]'
                                                        '/div/div[2]/div[1]/app-block-detail-card/div/div[2]/div[2]/'
                                                        'app-feature-item[4]/div/div[2]/div/a'))).text
                    bidding = WebDriverWait(browser, 1).until(
                        EC.presence_of_element_located((By.XPATH,
                                                        '/html/body/app-root/div/main/app-page-sales-item/'
                                                        'app-catalog-detail/div/'
                                                        f'div[{div_tag_number}]/div/div[2]/'
                                                        f'div[1]/app-block-detail-card/div/div[2]/div[2]/'
                                                        'app-feature-item[5]/div/div[2]/div/a')))
                except Exception:
                    platform = WebDriverWait(browser, 1).until(
                        EC.presence_of_element_located((By.XPATH,
                                                        '/html/body/app-root/div/main/app-page-sales-item/'
                                                        'app-catalog-detail/div/'
                                                        f'div[{div_tag_number}]'
                                                        '/div/div[2]/div[1]/app-block-detail-card/div/div[2]/div[2]/'
                                                        'app-feature-item[5]/div/div[2]/div/a'))).text
                    bidding = WebDriverWait(browser, 1).until(
                        EC.presence_of_element_located((By.XPATH,
                                                        '/html/body/app-root/div/main/app-page-sales-item/'
                                                        'app-catalog-detail/div/'
                                                        f'div[{div_tag_number}]/div/div[2]/'
                                                        f'div[1]/app-block-detail-card/div/div[2]/div[2]/'
                                                        'app-feature-item[6]/div/div[2]/div/a')))

                bidding_number = int(bidding.text)
                platform_link = bidding.get_attribute("href")
            except Exception:
                try:
                    platform = WebDriverWait(browser, 1).until(
                        EC.presence_of_element_located((By.XPATH,
                                                        '/html/body/app-root/div/main/app-page-sales-item/'
                                                        'app-catalog-detail/div/'
                                                        f'div[{div_tag_number}]/div/div[2]/'
                                                        f'div[1]/app-block-detail-card/div/div[2]/div[2]/'
                                                        'app-feature-item[4]/div/div[2]/div/a'))).text
                except Exception:
                    platform = WebDriverWait(browser, 1).until(
                        EC.presence_of_element_located((By.XPATH,
                                                        '/html/body/app-root/div/main/app-page-sales-item/'
                                                        'app-catalog-detail/div/'
                                                        f'div[{div_tag_number}]/div/div[2]/'
                                                        f'div[1]/app-block-detail-card/div/div[2]/div[2]/'
                                                        'app-feature-item[5]/div/div[2]/div/a'))).text
                bidding_number = "Не найден"
                platform_link = ""
            # Торги

            app_log.info("Парсинг торгов...")

            bidding_list = list()
            click_button(
                f"/html/body/app-root/div/main/app-page-sales-item/"
                f"app-catalog-detail/div/div[{div_tag_number}]/div/div[2]/div[1]/"
                "app-block-detail-card/div/div[2]/div[2]/app-feature-item[2]/div[1]/div/a")
            sleep(1)

            bid_div = browser.find_elements(By.XPATH, "/html/body/app-root/div/main/app-page-sales-item/"
                                                      f"app-catalog-detail/div/div[{div_tag_number}]/div/div[2]/div[1]/"
                                                      "app-block-detail-card/div/div[2]/div[2]/app-feature-item[2]/"
                                                      "div[2]/div[2]/div/div")
            for div in bid_div:
                info_bid = div.text.split("\n")
                for index in range(0, len(info_bid), 3):
                    bidding_list.append(info_bid[index:index + 3])
                break

            # Договор

            app_log.info("Успешно! Парсинг договоров...")
            treaty_list = list()

            for i in range(1, 100):
                try:
                    click_button('/html/body/app-root/div/main/app-page-sales-item/app-catalog-detail/div/'
                                 f'div[{div_tag_number}]/div/div[1]/div/div[2]/div[{i}]/div[1]/button')

                except Exception:
                    break
                app_log.debug(f"Договор {i}...")
                sleep(2)
                treaty_number = WebDriverWait(browser, 1).until(
                    EC.presence_of_element_located((By.XPATH,
                                                    '/html/body/app-root/div/main/app-page-sales-item/app-catalog'
                                                    '-detail/div/'
                                                    f'div[{div_tag_number}]/div/div[1]/'
                                                    f'div/div[3]/app-block-feature-group/div/div[2]/div/'
                                                    'app-block-feature-list/div[7]/div[2]/app-block-feature-value/'
                                                    'span'))).text
                loan_type = WebDriverWait(browser, 1).until(
                    EC.presence_of_element_located((By.XPATH,
                                                    '/html/body/app-root/div/main/app-page-sales-item/'
                                                    'app-catalog-detail/div/'
                                                    f'div[{div_tag_number}]/div/div[1]/div/'
                                                    f'div[3]/app-block-feature-group/div/div[2]/div/'
                                                    'app-block-feature-list/div[1]/div[2]/'
                                                    'app-block-feature-value/span'))).text
                treaty_date = WebDriverWait(browser, 1).until(
                    EC.presence_of_element_located((By.XPATH,
                                                    '/html/body/app-root/div/main/app-page-sales-item/'
                                                    'app-catalog-detail/div/'
                                                    f'div[{div_tag_number}]/div/div[1]/div/div[3]/'
                                                    f'app-block-feature-group/div/div[2]/div/'
                                                    'app-block-feature-list/div[2]/div[2]/'
                                                    'app-block-feature-value/span'))).text
                interest_rate = WebDriverWait(browser, 1).until(
                    EC.presence_of_element_located((By.XPATH,
                                                    '/html/body/app-root/div/main/app-page-sales-item/'
                                                    'app-catalog-detail/div/'
                                                    f'div[{div_tag_number}]/div/div[1]/div/div[3]/'
                                                    f'app-block-feature-group/div/div[2]/div/'
                                                    'app-block-feature-list/div[3]/div[2]/'
                                                    'app-block-feature-value/span'))).text
                interest_rate = float(interest_rate.replace(',', '.'))
                is_maturity_date = WebDriverWait(browser, 1).until(
                    EC.presence_of_element_located((By.XPATH,
                                                    '/html/body/app-root/div/main/app-page-sales-item/'
                                                    f'app-catalog-detail/div/div[{div_tag_number}]/'
                                                    f'div/div[1]/div/div[3]/app-block-feature-group/div/div[2]/div/'
                                                    'app-block-feature-list/div[4]/div[1]/span'))).text
                count_elem = 4
                if is_maturity_date == "Дата погашения":
                    count_elem += 1
                repayment_method = WebDriverWait(browser, 1).until(
                    EC.presence_of_element_located((By.XPATH,
                                                    '/html/body/app-root/div/main/app-page-sales-item/'
                                                    'app-catalog-detail/div/'
                                                    f'div[{div_tag_number}]/div/div[1]/div/'
                                                    f'div[3]/app-block-feature-group/div/div[2]/div/'
                                                    f'app-block-feature-list/div[{count_elem}]/div[2]/'
                                                    'app-block-feature-value/span'))).text
                is_estate = WebDriverWait(browser, 1).until(
                    EC.presence_of_element_located((By.XPATH,
                                                    '/html/body/app-root/div/main/app-page-sales-item/'
                                                    'app-catalog-detail/div/'
                                                    f'div[{div_tag_number}]/div/div[1]/'
                                                    f'div/div[3]/app-block-feature-group/div/div[2]/div/'
                                                    f'app-block-feature-list/div[{count_elem + 1}]/div[2]/'
                                                    'app-block-feature-value/span'))).text
                date_balance = WebDriverWait(browser, 1).until(
                    EC.presence_of_element_located((By.XPATH,
                                                    '/html/body/app-root/div/main/app-page-sales-item/'
                                                    'app-catalog-detail/div/'
                                                    f'div[{div_tag_number}]/div/div[1]/'
                                                    f'div/div[3]/app-block-feature-group/div/div[2]/div/'
                                                    f'app-block-feature-list/div[{count_elem + 4}]/div[2]/'
                                                    'app-block-feature-value/span'))).text
                debt_amount = WebDriverWait(browser, 1).until(
                    EC.presence_of_element_located((By.XPATH,
                                                    '/html/body/app-root/div/main/app-page-sales-item/'
                                                    'app-catalog-detail/div/'
                                                    f'div[{div_tag_number}]/div/div[1]/div/div[3]/'
                                                    'app-block-feature-group/div/div[2]/div/'
                                                    f'app-block-feature-list/div[{count_elem + 5}]/div[2]/'
                                                    'app-block-feature-value/span'))).text
                debt_amount = float(debt_amount.replace(',', '.'))
                try:
                    days_overdue = int(WebDriverWait(browser, 1).until(
                        EC.presence_of_element_located((By.XPATH,
                                                        '/html/body/app-root/div/main/app-page-sales-item/'
                                                        'app-catalog-detail/div/'
                                                        f'div[{div_tag_number}]/div/div[1]/'
                                                        f'div/div[3]/app-block-feature-group/div/div[2]/div/'
                                                        f'app-block-feature-list/div[{count_elem + 6}]/div[2]/'
                                                        'app-block-feature-value/span'))).text)
                except Exception:
                    days_overdue = "НЕ НАЙДЕНО"
                    count_elem -= 1
                is_court_rulings = WebDriverWait(browser, 1).until(
                    EC.presence_of_element_located((By.XPATH,
                                                    '/html/body/app-root/div/main/app-page-sales-item/'
                                                    'app-catalog-detail/div/'
                                                    f'div[{div_tag_number}]/div/div[1]/div/div[3]/'
                                                    f'app-block-feature-group/div/div[2]/div/'
                                                    f'app-block-feature-list/div[{count_elem + 7}]/div[2]/'
                                                    'app-block-feature-value/span'))).text
                try:
                    maturity_date = WebDriverWait(browser, 1).until(
                        EC.presence_of_element_located((By.XPATH,
                                                        '/html/body/app-root/div/main/app-page-sales-item/'
                                                        'app-catalog-detail/div/'
                                                        f'div[{div_tag_number}]/div/'
                                                        f'div[1]/div/div[3]/app-block-feature-group/div/div[2]/div/'
                                                        f'app-block-feature-list/div[{count_elem + 8}]/div[2]/'
                                                        'app-block-feature-value/span'))).text
                except Exception:
                    maturity_date = "НЕ НАЙДЕНО"
                real_estate_list = list()
                if is_estate == "Да":
                    dict_of_char = dict()
                    for j in range(1, 100):
                        dict_of_char = dict()
                        try:
                            click_button('/html/body/app-root/div/main/app-page-sales-item/app-catalog-detail/'
                                         f'div/div[{div_tag_number}]/div/'
                                         'div[1]/div/div[3]/app-block-feature-group/div/div[2]/div/div/div[2]/'
                                         f'app-block-collaterals/div/div/div[{j}]/a')
                            sleep(2)
                            for k in range(1, 100):
                                try:
                                    key_name = WebDriverWait(browser, 1).until(
                                        EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div/'
                                                                                  'app-modal-collaterals/'
                                                                                  'app-popup-default/'
                                                                                  'div/div/div[2]/div[1]/div[2]/'
                                                                                  'app-block-feature-group/div/div[2]/'
                                                                                  f'div/app-block-feature-list/'
                                                                                  f'div[{k}]/'
                                                                                  'div[1]/span'))).text

                                    value_name = WebDriverWait(browser, 1).until(
                                        EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div/'
                                                                                  'app-modal-collaterals/'
                                                                                  'app-popup-default/'
                                                                                  'div/div/div[2]/div[1]/div[2]/'
                                                                                  'app-block-feature-group/div/div[2]/'
                                                                                  f'div/app-block-feature-list/'
                                                                                  f'div[{k}]/'
                                                                                  f'div[2]/'
                                                                                  f'app-block-feature-value/'
                                                                                  f'span'))).text
                                    if key_name in KEYS_LIST:
                                        dict_of_char[key_name] = value_name
                                except Exception:
                                    cur_button = WebDriverWait(browser, 4).until(
                                        EC.presence_of_element_located(
                                            (By.XPATH, '/html/body/div[2]/div/div/app-modal-collaterals/'
                                                       'app-popup-default/div/div/div[2]/div[2]/app-icon'))
                                    )
                                    browser.execute_script("arguments[0].click();", cur_button)
                                    browser.execute_script(f"window.scrollTo(0, 1000)")
                                    break
                        except Exception as e:
                            app_log.info(f"Найдено обеспечений: {j - 1}")
                            break
                        real_estate_list.append(dict_of_char)

                treaty_list.append([treaty_number, loan_type, treaty_date, interest_rate, repayment_method, is_estate,
                                    date_balance, debt_amount, days_overdue, is_court_rulings, maturity_date,
                                    real_estate_list])
            if platform_link and platform != "АО «АГЗРТ»":
                app_log.info(f"Успешно! Начинаю парсить на catalog.lot-online.ru ({platform_link})")
                # Данные из карточки лота на сайте catalog.lot-online.ru
                if "catalog.lot-online.ru" not in platform_link:
                    app_log.error(f"Внимание! ссылка {platform_link} не является catalog.lot-online.ru!")
                    maturity_date, address = "", ""
                else:
                    browser.get(platform_link)
                    try:
                        link_to_lot = WebDriverWait(browser, 1).until(
                            EC.presence_of_element_located((By.XPATH,
                                                            '/html/body/div[1]/div[4]/div[4]/div/div[2]/div/div[2]/'
                                                            'div[2]/div/div/div[3]/div/div[2]/div[1]/div/form/'
                                                            'div[1]/div[1]/div/div[1]/'
                                                            'div/div[1]/div/a'))).get_attribute("href")
                    except Exception:
                        try:
                            link_to_lot = WebDriverWait(browser, 1).until(
                                EC.presence_of_element_located((By.XPATH,
                                                                '/html/body/div[1]/div[4]/div[4]/div/div[2]/div/div[2]/'
                                                                'div[2]/div/div/div[3]/div/div[2]/div[1]/div/form/'
                                                                'div[2]/div[1]/'
                                                                'bdi/a'))).get_attribute("href")
                        except Exception:
                            link_to_lot = WebDriverWait(browser, 1).until(
                                EC.presence_of_element_located((By.XPATH,
                                                                '/html/body/div[1]/div[4]/div[4]/div/div[2]/div/'
                                                                'div[2]/div[2]/div/div/div[3]/div/div[3]/div[1]/'
                                                                'div/form/div[2]/div[1]/'
                                                                'bdi/a'))).get_attribute("href")
                    browser.get(link_to_lot)
                    try:
                        click_button('//*[@id="ui-id-1"]')
                    except Exception:
                        pass
                    try:
                        maturity_date = WebDriverWait(browser, 1).until(
                            EC.presence_of_element_located((By.XPATH,
                                                            '//*[@id="ui-id-2"]/div[6]/span[2]'))).text
                    except Exception:
                        maturity_date = "Не найдено"
                    address = WebDriverWait(browser, 1).until(
                        EC.presence_of_element_located((By.XPATH,
                                                        '//*[@id="tygh_main_container"]/div[4]/div/div[2]/div/'
                                                        'div[1]/div/div[2]/'
                                                        'div[1]/div[2]/form[1]/div[8]/dl/div[6]/dd'))).text
            else:
                app_log.warning("Не найдена ссылка на catalog.lot-online.ru!")
                maturity_date, address = "Не найдено", "Не найдено"

            app_log.info(f"Формирую excel файл {num_page}.xlsx ...")
            writer = pd.ExcelWriter(f'{num_page}.xlsx', engine='xlsxwriter')
            workbook = writer.book
            worksheet = workbook.add_worksheet(num_page)
            writer.sheets[num_page] = worksheet

            base_df = pd.DataFrame({"Наименование": [
                "Лот на АСВ", "Название лота", "Номер лота", "Балансовая стоимость, руб.", "Тип кредита",
                "Вид имущества",
                "Регион",
                "ФО/Агентство", "Дата публикации", "Ссылка на публикацию", "ЭП", "Ссылка на ЭП", "Номер торга"
            ],
                "Значение": [
                    url_page, lot_name, lot_number, price, type_of_credit, type_of_property, region, agency,
                    data_published, public_link, platform, platform_link, bidding_number
                ]})

            start_bid, stop_bid, price_bid, type_bid = list(), list(), list(), list()
            for bid in bidding_list:
                type_bid.append(bid[0])
                price_bid.append(float("".join(bid[1][:-1].split("\u2009")).replace(",", ".")))
                try:
                    start_bid_elem, stop_bid_elem = bid[2].split(" - ")
                except ValueError:
                    start_bid_elem, stop_bid_elem = bid[2].split(" - ")[0], "Не найден"
                start_bid.append(start_bid_elem)
                stop_bid.append(stop_bid_elem)
            bidding_df = pd.DataFrame({
                "Начало торгов": start_bid,
                "Окончание торгов": stop_bid,
                "Цена": price_bid,
                "Тип торгов": type_bid
            })

            base_df.to_excel(writer, sheet_name=num_page, startrow=0, startcol=0, index=False)
            bidding_df.to_excel(writer, sheet_name=num_page, startrow=0, startcol=3)

            worksheet.set_column("A:B", 60)
            glob_index = 0
            for i_treaty, treaty in enumerate(treaty_list):
                df_treaty = pd.DataFrame({
                    "Наименование": ["Номер договора", "Тип кредита", "Дата заключения договора",
                                     "Процентная ставка, %",
                                     "Способ погашения задолженности", "Наличие обеспечения", "Дата баланса",
                                     "Размер задолженности на текущую дату", "Кол-во дней просроченного платежа",
                                     "Наличие судебных разбирательств", "Дата последнего погашения", "Поручительство"],
                    "Значение": [treaty[0], treaty[1], treaty[2], treaty[3], treaty[4], treaty[5],
                                 treaty[6], treaty[7], treaty[8], treaty[9], treaty[10], ""]
                })
                df_treaty.to_excel(writer, sheet_name=num_page, startrow=15 + 16 * i_treaty + 1, startcol=0,
                                   index=False)

                cur_estate_list = treaty[11]
                for index, estate in enumerate(cur_estate_list):
                    cur_keys = estate.keys()
                    result_data = {
                        "Тип": estate["Тип"] if "Тип" in cur_keys else "",
                        "Площадь": estate["Площадь"] if "Площадь" in cur_keys else "",
                        "Количество комнат": estate["Количество комнат"] if "Количество комнат" in cur_keys else "",
                        "№ этажа": estate["№ этажа"] if "№ этажа" in cur_keys else "",
                        "Количество этажей / этажность": estate["Количество этажей"] if "Количество этажей" in
                                                                                        cur_keys else
                        estate["Этажность"] if "Этажность" in cur_keys else "",
                        "Тип дома": estate["Тип дома"] if "Тип дома" in cur_keys else "",
                        "Материал стен": estate["Материал стен"] if "Материал стен" in cur_keys else "",
                        "Категория": estate["Категория"] if "Категория" in cur_keys else "",
                        "Разрешенное использование": estate["Разрешенное использование"] if "Разрешенное использование"
                                                                                            in
                                                                                            cur_keys else "",
                        "Кадастровый (условный) номер": estate["Кадастровый (условный) номер"] if
                        "Кадастровый (условный) номер"
                        in cur_keys else "",
                        "Адрес": estate["Адрес"] if "Адрес" in cur_keys else "",
                        "Год постройки": estate["Год постройки"] if "Год постройки" in cur_keys else "",
                        "Коммуникации и их характеристики": estate["Коммуникации и их характеристики"] if
                        "Коммуникации и их характеристики" in cur_keys else "",
                        "Ограничения и обременения": estate[
                            "Ограничения и обременения"] if "Ограничения и обременения" in
                                                            cur_keys else ""
                    }
                    estate_df = pd.DataFrame({
                        "Наименование": result_data.keys(),
                        "Значение": result_data.values()
                    })
                    glob_index = 3 + index * 3
                    estate_df.to_excel(writer, sheet_name=num_page, startrow=16 + 16 * i_treaty, startcol=glob_index,
                                       index=False)

            worksheet.set_column("D:H", 25)
            worksheet.set_column(3, glob_index + 1, 25)
            writer.close()
            app_log.info("Парсинг успешно завершен!")
        except Exception as e:
            app_log.critical(e)
    app_log.info("Закрываю процессы...")
    browser.close()
    os.system("taskkill /f /IM chrome.exe >nul 2>&1")
    os.system("taskkill /f /IM chromedriver.exe >nul 2>&1")
    app_log.info("Закончил обработку всех ссылок!")
