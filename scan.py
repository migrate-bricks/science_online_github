# -*- coding: utf-8 -*-
# before using this script, please refer to d = u2.connect("AUE66HL7XWIVJRSS") Change it based on 'adb devices'

import json
import logging
import os
import random
import re
import sys
import time
from datetime import datetime

import colorlog
import openpyxl
import openpyxl.utils
import uiautomator2 as u2
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

console_handler = logging.StreamHandler()
console_handler.setLevel(logging.DEBUG)

formatter = colorlog.ColoredFormatter(
    fmt='%(log_color)s %(asctime)s %(levelname)s:%(message)s%(reset)s',
    datefmt="%Y-%m-%d %H:%M:%S",
    log_colors={'DEBUG': 'cyan', 'INFO': 'green', 'WARNING': 'yellow', 'ERROR': 'red', 'CRITICAL': 'red,bg_white'}
)

console_handler.setFormatter(formatter)
logger.addHandler(console_handler)

package_name = "com.taobao.idlefish"
browser_package_name = "com.android.browser"
activity_name = ".maincontainer.activity.MainActivity"


class TimeUtil:
    @staticmethod
    def random_sleep(random_start=2, random_end=5):
        wait_time = random.randint(random_start, random_end)
        time.sleep(wait_time)

    @staticmethod
    def sleep(secs: float) -> None:
        time.sleep(secs)

    @staticmethod
    def curr_date() -> str:
        return datetime.now().strftime("%Y-%m-%d")


def read_excel(excel_path: str, header_row: int) -> list:
    # Open the Excel file
    wb = openpyxl.load_workbook(excel_path, read_only=True, keep_vba=False, data_only=True, keep_links=False, rich_text=False)

    # Select the worksheet you want to read
    ws = wb.active

    # Get the column names
    columns = [cell.value for cell in next(ws.iter_rows(header_row))]  # second row is header

    # Read the data
    results = []
    for row in ws.iter_rows(min_row=header_row+1, values_only=True):
        if row[0] is None:
            break
        results.append({columns[i]: value for i, value in enumerate(row)})
    return results


def append_excel(data_list: list, output_file: str) -> None:
    if len(data_list) <= 0:
        return False

    if not os.path.exists(output_file):
        save_excel(data_list, output_file)
    else:
        wb = openpyxl.load_workbook(output_file)
        ws = wb.active
        ws.append(data_list.values())
        wb.save(output_file)


def save_excel(data_list: list, output_file: str) -> None:
    if len(data_list) <= 0:
        return False

    # Create file if not exist
    if not os.path.exists(os.path.dirname(output_file)):
        os.makedirs(os.path.dirname(output_file))

    # Open a new workbook
    wb = Workbook()
    ws = wb.active

    # Write the column header and apply formatting, set the width of each column based on the length of its name
    headers = list(data_list[0].keys())
    for i, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=i, value=header)
        cell.fill = PatternFill(start_color='00c5d9f1', end_color='00c5d9f1', fill_type='solid')
        ws.column_dimensions[cell.column_letter].width = max(10, len(header))
        # ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = max(10, len(header))

    # Write the data
    for row, row_data in enumerate(data_list, 2):
        for column, value in enumerate(row_data.values(), 1):
            ws.cell(row=row, column=column, value=value)

    # Determine the maximum number of columns
    max_column = max(len(row) for row in ws.rows)

    # Set the autofilter for all columns
    range_string = f"A1:{openpyxl.utils.get_column_letter(max_column)}{ws.max_row}"
    ws.auto_filter.ref = range_string
    # Freeze the top row
    ws.freeze_panes = ws['A2']
    wb.save(output_file)


def get_save_folder() -> str:
    date = TimeUtil.curr_date()
    return os.path.join(os.getcwd(), 'save', date)


def get_save_path(filename: str) -> str:
    return os.path.join(get_save_folder(), filename)


def swipe_up():
    d.swipe_ext('up', 0.9)


def open_page_by_keyword(search_keyword: str):
    d.app_start(package_name, activity_name, wait=True)
    d(resourceId="com.taobao.idlefish:id/search_bar_layout").must_wait()
    d(resourceId="com.taobao.idlefish:id/search_bar_layout").click()
    d.send_keys(search_keyword, clear=True)
    d.press('enter')


def open_page_by_url(url: str):
    d.app_stop(browser_package_name)
    d.app_stop(package_name)
    d.app_start(browser_package_name, wait=True)
    d(resourceId="com.android.browser:id/search_hint").must_wait()
    d(resourceId="com.android.browser:id/search_hint").click_exists()
    d(resourceId="com.android.browser:id/url").must_wait()
    d(resourceId="com.android.browser:id/url").set_text(url)
    d.press('enter')
    d(textContains='允许').wait(exists=True)
    d(textContains='允许').click(timeout=30)


def get_platform_price(s: str) -> float:
    # Platform price format is different from Store price format
    match = re.search(r'¥(\d+\.?\d*)', s)
    if match:
        price = match.group(1)
        return float(price)


def get_store_price(s: str) -> float:
    # Platform price format is different from Store price format
    match = re.search(r'商品价格(\d+\.?\d*)', s)
    if match:
        price = match.group(1)
        return float(price)
    return 0


def get_wanted(s: str) -> float:
    match = re.search(r'(\d+\.?\d*)人想要', s)
    if match:
        price = match.group(1)
        return float(price)
    return 0


def get_min_price_but_greater_than_one(results: list) -> float:
    # Get the smallest price but skip those < 1, it's meaningless to do <1 biz
    min_price = sys.maxsize
    for item in results:
        if 1 <= item['price'] and item['price'] < min_price:
            min_price = item['price']
    return min_price


def get_comebine_prices(results: list) -> str:
    prices = [str(item['price']) for item in results]
    return ",".join(sorted(prices))


def get_store_result_by_key(store_results: list, key: str):
    for store in store_results:
        if key.lower() in store['title'].lower():
            return store
    return {}


def clean_platform_text(text: str) -> str:
    return text.replace('\n', '')


def clean_text(text: str) -> str:
    return text.replace('\n', '@')


def main_complete():
    d.set_fastinput_ime(False)


def should_scan_store(store_excel_file_path: str) -> bool:
    if not os.path.exists(store_excel_file_path):
        return True
    store_name = os.path.splitext(os.path.basename(store_excel_file_path))[0]
    print(f'{store_name} results are already exists, path: {store_excel_file_path}')
    overwrite = input('【Overwrite? 1.Yes, 2.No】: ')
    return (overwrite == '1')


def scan_platform(idx: int, search_keywords: str, must_include_word: str, max_scroll_page: int, scroll_page_timeout: float):
    try:
        logger.info(d.info)
        logger.info(f"Retrieving products information for 【{search_keywords}】...")
        results = []
        for search_keyword in search_keywords:
            open_page_by_keyword(search_keyword)
            for i in range(max_scroll_page):
                logger.info(f"Scrolling to idx: {idx}, keyword: {search_keyword}, include: {must_include_word} [{i}/{max_scroll_page}] page...")
                TimeUtil.sleep(scroll_page_timeout)
                view_list = d.xpath('//android.widget.ScrollView//android.view.View').all()
                if len(view_list) > 0:
                    for el in view_list:
                        if len(el.elem.getchildren()) > 0:
                            el_description = clean_platform_text(str(el.attrib['content-desc']))
                            for child in el.elem.getchildren():
                                el_description = f"{el_description}&{clean_platform_text(str(child.attrib['content-desc']))}"  # combine el_description
                            if must_include_word.lower() in el_description.lower():
                                price = get_platform_price(el_description)
                                wanted = get_wanted(el_description)
                                # skip duplicated item
                                if price is not None and price != '' and not any(d['title'] == el_description for d in results):
                                    logger.info(f"【{len(results)+1}】-description:{el_description}, price:{price}, wanted:{wanted}")
                                    results.append({'title': el_description, 'price': price, 'wanted': wanted})
                if d(descriptionContains='到底了').exists:  # alread on the end of the page
                    break
                swipe_up()
        return results
    except Exception as e:
        print(e)
        logger.error("Program runs Error:" + str(e.args[0]))
    finally:
        main_complete()
        print("Execution Completed!")


def scan_store(store_name: str, must_include_word: str, max_scroll_page: int, scroll_page_timeout: float):
    try:
        logger.info(d.info)
        logger.info(f"Retrieving products information for【{store_name} ...")
        results = []
        for idx in range(max_scroll_page):
            logger.info(f"Scrolling to store: {store_name} [{idx}/{max_scroll_page}] page...")
            TimeUtil.sleep(scroll_page_timeout)
            view_list = d.xpath('//android.widget.ScrollView//android.view.View').all()
            if len(view_list) > 0:
                for el in view_list:
                    el_description = clean_text(str(el.attrib['content-desc']))
                    if must_include_word.lower() in el_description.lower():
                        price = get_store_price(el_description)
                        wanted = get_wanted(el_description)
                        if price is not None and price != '' and not any(d['title'] == el_description for d in results):  # Skip duplicated item
                            logger.info(f"【{len(results)+1}】-description:{el_description}, price:{price}, wanted:{wanted}")
                            results.append({'title': el_description, 'price': price, 'wanted': wanted})
            if d(descriptionContains='没有更多了').exists:  # Alread on the end of the page
                break
            swipe_up()
        return results
    except Exception as e:
        print(e)
        logger.error("Program runs Error:" + str(e.args[0]))
    finally:
        main_complete()
        print("Execution Completed!")


def load_store_results(store_name: str, store_homepage: str, store_must_include_word: str, store_max_scroll_page: int, store_scroll_page_timeout: float) -> list:
    store_excel_file_path = get_save_path(f"STORE_{store_name}.xlsx")
    if should_scan_store(store_excel_file_path):
        print(f'【Navigate to store: {store_name}, homepage: {store_homepage} ...】')
        open_page_by_url(store_homepage)
        print(f'【Scan store: {store_name}, homepage: {store_homepage}...')
        store_results = scan_store(store_name, store_must_include_word, store_max_scroll_page, store_scroll_page_timeout)
        store_sorted_results = sorted(store_results, key=lambda x: x['wanted'], reverse=True)
        print(f'【Save store {store_name} results, path: {store_excel_file_path}...')
        save_excel(store_sorted_results, store_excel_file_path)
    else:
        print(f'【Read store {store_name} results from existing path: {store_excel_file_path}')
        store_results = read_excel(store_excel_file_path, header_row=1)
    return store_results


def full_outer_join(store_results: list, delivery_settings: list):
    for store in store_results:  # Set must_include_word on store results
        for delivery in delivery_settings:
            if delivery['must_include_word'].lower() in store['title'].lower():
                store['must_include_word'] = delivery['must_include_word']
                break
    resultsDataFrame = pd.merge(pd.DataFrame(delivery_settings), pd.DataFrame(store_results), on='must_include_word', how='outer')
    return resultsDataFrame.to_dict(orient='records')


if __name__ == '__main__':
    print('【Please Make sure the uiautomator2 has connected to android device and setup the `correct android_device_addr` in scan_config.json】')
    
    with open('./scan_config.json', 'r', encoding='utf8') as fp:
        scan_config = json.load(fp)

        android_device_addr = scan_config['android_device_addr']
        scroll_page_timeout_second = scan_config['scroll_page_timeout_second']
        delivery_settings_path = scan_config['delivery_settings_path']

        platform_max_scroll_page = scan_config['scan_platform']['max_scroll_page']

        store_max_scroll_page = scan_config['scan_store']['max_scroll_page']
        store_must_include_word = scan_config['scan_store']['must_include_word']
        stores = scan_config['scan_store']['stores']

        # Define first store is the main store
        main_store_name = stores[0]['store_name']
        main_store_homepage = stores[0]['home_page']
        
        scan_type = input('【Scan type? 1.platform or 2.store】: ')
        d = u2.connect(android_device_addr)  # Change it based on 'adb devices'
        d.screen_on()
        if scan_type == '1':  # Scan platform
            print('【Scan platform Start】...')
            print(f'【Read Global delivery settings from path: {delivery_settings_path}】...')
            delivery_settings = read_excel(delivery_settings_path, header_row=2)  # Header:*配置名称|*外部编码|*附言（具体的发货内容填写在这里）|商品分类（选填）|自动发货开关（不填默认开启）|配置名称是否等于外部编码|附言是否包含外部编码|search_keywords|must_include_word|

            # Retrieve Main store results
            print('【Load Main store results for getting current price】...')
            main_store_results = load_store_results(main_store_name, main_store_homepage, store_must_include_word, store_max_scroll_page, scroll_page_timeout_second)

            summary_results = []
            for idx, delivery_setting in enumerate(delivery_settings):  # Excel settings is delivery_setting
                setting_search_keywords = delivery_setting['search_keywords'].split(',')
                setting_must_include_word = delivery_setting['must_include_word']  # Foreign key

                print(f'【Scan Platform by keyword: {setting_search_keywords}, include: {setting_must_include_word} ...')
                platform_results = scan_platform(idx, setting_search_keywords, setting_must_include_word, platform_max_scroll_page, scroll_page_timeout_second)

                platform_sorted_results = sorted(platform_results, key=lambda x: x['price'])
                platform_min_price = get_min_price_but_greater_than_one(platform_results)
                platform_combine_prices = get_comebine_prices(platform_results)

                # Find target main store result, get current title, wanted, price
                main_store_result = get_store_result_by_key(main_store_results, setting_must_include_word)
                if main_store_result is not None:
                    current_price = main_store_result.get('price')

                platform_excel_file_name = f"PLATFORM_{setting_must_include_word}_price_{current_price}_min_{platform_min_price}.xlsx"
                platform_excel_file_path = get_save_path(platform_excel_file_name)
                print(f'【{idx}_Save Platform results keywords:{setting_search_keywords}, include: {setting_must_include_word} path: {platform_excel_file_path}...')
                save_excel(platform_sorted_results, platform_excel_file_path)
                logger.info(f"【{idx}: Save Platform results completed, File path: {platform_excel_file_path}")

                summary_result = {**delivery_setting, **main_store_result, 'min_price': platform_min_price, 'combine_prices': platform_combine_prices}
                summary_results.append(summary_result)
                logger.info(f'【Summary result: {summary_result}')
                if (len(summary_results) % 30 == 0):
                    save_excel(summary_results, get_save_path(f"SUMMARY_PLATFORM_{len(summary_results)}.xlsx"))

            summary_excel_file_path = get_save_path("SUMMARY_PLATFORM.xlsx")
            save_excel(summary_results, summary_excel_file_path)
            logger.info(f"【Platform Summary Execution completed, File path: {summary_excel_file_path}")
        elif scan_type == '2':  # Scan Store
            print('【Scan store Start】...')
            for idx, store in enumerate(stores, start=1):
                print(f'*{idx}.{store['store_name']} {store['home_page']}')
            store_index = input('【Store Index?】: ')
            store_name = stores[int(store_index)-1]['store_name']
            store_homepage = stores[int(store_index)-1]['home_page']

            print(f'【Read Global delivery settings from path: {delivery_settings_path}】')
            delivery_settings = read_excel(delivery_settings_path, header_row=2)  # Header:*配置名称|*外部编码|*附言（具体的发货内容填写在这里）|商品分类（选填）|自动发货开关（不填默认开启）|配置名称是否等于外部编码|附言是否包含外部编码|search_keywords|must_include_word|

            print('【Load Store results】')
            store_results = load_store_results(store_name, store_homepage, store_must_include_word, store_max_scroll_page, scroll_page_timeout_second)

            mrege_results = full_outer_join(store_results, delivery_settings)

            summary_store_excel_file_path = get_save_path(f"SUMMARY_STORE_{store_name}.xlsx")
            print(f'【Save {store_name} STORE SUMMARY results, path: {summary_store_excel_file_path}...')
            save_excel(mrege_results, summary_store_excel_file_path)
