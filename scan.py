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
import uiautomator2 as u2
from openpyxl import Workbook

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
    def sleep(secs):
        time.sleep(secs)

    @staticmethod
    def curr_date():
        return datetime.now().strftime("%Y-%m-%d")


def read_excel(excel_path, header_row):
    # Open the Excel file
    wb = openpyxl.load_workbook(excel_path, read_only=True, keep_vba=False, data_only=True, keep_links=False, rich_text=False)

    # Select the worksheet you want to read
    sh = wb.active

    # Get the column names
    columns = [cell.value for cell in next(sh.iter_rows(header_row))]  # second row is header

    # Read the data
    results = []
    for row in sh.iter_rows(min_row=header_row+1, values_only=True):
        if row[0] is None:
            break
        results.append({columns[i]: value for i, value in enumerate(row)})
    return results


def append_excel(data_list, output_file):
    if len(data_list) <= 0:
        return False

    if not os.path.exists(output_file):
        save_excel(data_list, output_file)
    else:
        wb = openpyxl.load_workbook(output_file)
        sh = wb.active
        sh.append(data_list.values())
        wb.save(output_file)


def save_excel(data_list, output_file):
    if len(data_list) <= 0:
        return False

    # Create file if not exist
    if not os.path.exists(os.path.dirname(output_file)):
        os.makedirs(os.path.dirname(output_file))

    # Open a new workbook
    wb = Workbook()
    sh = wb.active

    # Write the column header
    headers = list(data_list[0].keys())
    for i, header in enumerate(headers, 1):
        sh.cell(row=1, column=i, value=header)

    # Write the data
    for row, row_data in enumerate(data_list, 2):
        for column, value in enumerate(row_data.values(), 1):
            sh.cell(row=row, column=column, value=value)
    wb.save(output_file)


def get_save_folder():
    date = TimeUtil.curr_date()
    return os.path.join(os.getcwd(), 'save', date)


def get_save_path(filename):
    return os.path.join(get_save_folder(), filename)


def swipe_up():
    d.swipe_ext('up', 0.9)


def open_page_by_keyword(search_keyword):
    d.app_start(package_name, activity_name, wait=True)
    d(resourceId="com.taobao.idlefish:id/search_bar_layout").must_wait()
    d(resourceId="com.taobao.idlefish:id/search_bar_layout").click()
    d.send_keys(search_keyword, clear=True)
    d.press('enter')


def open_page_by_url(url):
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


def get_platform_price(s):
    # Platform price format is different from Store price format
    match = re.search(r'¥(\d+\.?\d*)', s)
    if match:
        price = match.group(1)
        return float(price)


def get_store_price(s):
    # Platform price format is different from Store price format
    match = re.search(r'商品价格(\d+\.?\d*)', s)
    if match:
        price = match.group(1)
        return float(price)
    return 0


def get_wanted(s):
    match = re.search(r'(\d+\.?\d*)人想要', s)
    if match:
        price = match.group(1)
        return float(price)
    return 0


def get_min_price_but_greater_than_one(results):
    # Get the smallest price but skip those < 1, it's meaningless to do <1 biz
    min_price = sys.maxsize
    for item in results:
        if 1 <= item['price'] and item['price'] < min_price:
            min_price = item['price']
    return min_price


def get_comebine_prices(results):
    prices = [str(item['price']) for item in results]
    return ",".join(sorted(prices))


def get_store_result_by_key(store_results, key):
    for store in store_results:
        if key.lower() in store['title'].lower():
            return store
    return {}


def clean_platform_text(text):
    return text.replace('\n', '')


def clean_text(text):
    return text.replace('\n', '@')


def main_complete():
    d.set_fastinput_ime(False)


def should_scan_store(store_excel_file_path):
    if not os.path.exists(store_excel_file_path):
        return True
    overwrite = input(f'Main store results are already exists, do you want to overwrite Yes(1), No(0) ? Path: {store_excel_file_path}')
    return (overwrite == '1')


def scan_platform(idx, search_keywords, must_include_word, max_scroll_page):
    try:
        logger.info(d.info)
        logger.info(f"Retrieving products information for 【{search_keywords}】...")
        results = []
        for search_keyword in search_keywords:
            open_page_by_keyword(search_keyword)
            for i in range(max_scroll_page):
                logger.info(f"Scrolling to {idx}.{search_keyword} [{i}/{max_scroll_page}] page...")
                TimeUtil.sleep(2)
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


def scan_store(store_name, must_include_word, max_scroll_page):
    try:
        logger.info(d.info)
        logger.info(f"Retrieving products information for【{store_name} ...")
        results = []
        for i in range(max_scroll_page):
            logger.info(f"Scrolling to [{i}/{max_scroll_page}] page...")
            TimeUtil.sleep(2)
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


if __name__ == '__main__':
    with open('./scan_config.json', 'r', encoding='utf8') as fp:
        scan_config = json.load(fp)

        android_device_addr = scan_config['android_device_addr']
        delivery_settings_path = scan_config['delivery_settings_path']

        platform_max_scroll_page = scan_config['scan_platform']['max_scroll_page']

        store_max_scroll_page = scan_config['scan_store']['max_scroll_page']
        store_must_include_word = scan_config['scan_store']['must_include_word']
        stores = scan_config['scan_store']['stores']

        # Define first store is the main store
        main_store_name = stores[0]['store_name']
        main_store_homepage = stores[0]['home_page']

        d = u2.connect(android_device_addr)  # TODO: Change it based on 'adb devices' "AUE66HL7XWIVJRSS"

        scan_type = input('Scan: 1.platform or 2.store, press Enter will scan 1.platform by default? ')
        if scan_type == '1' or scan_type == '':  # Scan platform
            print('【Scan platform start, and load main store firstly for getting current price】...')
            # Retrieve Main store results
            main_store_excel_file_path = get_save_path(f"STORE_{main_store_name}.xlsx")
            if should_scan_store(main_store_excel_file_path):
                print('【Navigate to Main store homepage...】')
                open_page_by_url(main_store_homepage)
                print('【Scan Main store home page...')
                main_store_results = scan_store(store_name=main_store_name, must_include_word=store_must_include_word, max_scroll_page=store_max_scroll_page)
                main_store_sorted_results = sorted(main_store_results, key=lambda x: x['wanted'], reverse=True)
                print(f'Save Main store results, path: {main_store_excel_file_path}...')
                save_excel(main_store_sorted_results, main_store_excel_file_path)
            else:
                print(f'【Read Main store results from existing path: {main_store_excel_file_path}')
                main_store_results = read_excel(main_store_excel_file_path, header_row=1)

            print(f'【Read Global delivery settings from path: {delivery_settings_path}】')
            delivery_settings = read_excel(delivery_settings_path, header_row=2)  # Header:*配置名称|*外部编码|*附言（具体的发货内容填写在这里）|商品分类（选填）|自动发货开关（不填默认开启）|配置名称是否等于外部编码|附言是否包含外部编码|search_keywords|must_include_word|

            summary_results = []
            for idx, delivery_setting in enumerate(delivery_settings):  # Excel settings is delivery_setting
                setting_search_keywords = delivery_setting['search_keywords'].split(',')
                setting_must_include_word = delivery_setting['must_include_word']

                print(f'【Scan Platform by keyword: {setting_search_keywords}, include: {setting_must_include_word} ...')
                platform_results = scan_platform(idx=idx,
                                                 search_keywords=setting_search_keywords,
                                                 must_include_word=setting_must_include_word,
                                                 max_scroll_page=platform_max_scroll_page)

                platform_sorted_results = sorted(platform_results, key=lambda x: x['price'])
                min_price = get_min_price_but_greater_than_one(platform_results)
                combine_prices = get_comebine_prices(platform_results)

                # Find target main store result, get current title, wanted, price
                main_store_result = get_store_result_by_key(main_store_results, setting_must_include_word)
                if main_store_result is not None:
                    current_price = main_store_result.get('price')

                platform_excel_file_name = f"PLATFORM_{setting_must_include_word}_price-{current_price}_min-{min_price}.xlsx"
                platform_excel_file_path = get_save_path(platform_excel_file_name)
                print(f'【{idx}.Save Platform results include: {setting_must_include_word} path: {platform_excel_file_path}...')
                save_excel(platform_sorted_results, platform_excel_file_path)
                logger.info(f"【{idx}.Save Platform results completed, File path: {platform_excel_file_path}")

                summary_result = {**delivery_setting, **main_store_result, 'min_price': min_price, 'combine_prices': combine_prices}
                summary_results.append(summary_result)
                logger.info(f'【Summary result】: {summary_result}')

            summary_excel_file_path = get_save_path("SUMMARY_PLATFORM.xlsx")
            save_excel(summary_results, summary_excel_file_path)
            logger.info(f"【Platform Summary】 Execution completed, File path: {summary_excel_file_path}")
        elif scan_type == '2':  # Scan platform
            print('Scan store...')

    # with open('./scan_platform_config.json', 'r', encoding='utf8') as fp:
    #     platform_config = json.load(fp)

    #     android_device_addr = platform_config["android_device_addr"]
    #     max_scroll_page = platform_config['max_scroll_page']
    #     delivery_settings_path = platform_config['delivery_settings_path']
    #     # searchs = read_delivery_settings(delivery_settings_path)
    #     delivery_settings = read_excel(delivery_settings_path, 2)

    #     d = u2.connect(android_device_addr)  # TODO: Change it based on 'adb devices' "AUE66HL7XWIVJRSS"

    #     for delivery_setting in delivery_settings:
    #         setting_search_keywords = delivery_setting['search_keywords']
    #         setting_must_include_word = delivery_setting['must_include_word']
    #         platform_results = scan_platform(search_keywords=setting_search_keywords, must_include_word=setting_must_include_word, max_scroll_page=max_scroll_page)

    #         platform_sorted_results = sorted(platform_results, key=lambda x: x['price'])
    #         min_price = get_min_price_but_greater_than_one(platform_results)
    #         excel_file_name = f"{setting_must_include_word}-{min_price}.xlsx"
    #         save_excel(platform_sorted_results, get_save_path(excel_file_name))
    #         # output_file = to_excel(results, must_include_word)
    #         logger.info(f"Execution completed, file path: {excel_file_name}")

    # with open('./scan_store_config.json', 'r', encoding='utf8') as fp:
    #     store_config = json.load(fp)
    #     d = u2.connect(store_config["android_device_addr"])  # TODO: Change it based on 'adb devices' "AUE66HL7XWIVJRSS"
    #     setting_must_include_word = store_config['must_include_word']
    #     max_scroll_page = store_config['max_scroll_page']

    #     print('【All available stores】:')
    #     for idx, store in enumerate(store_config['stores']):
    #         print(f"【{idx}】, {store['store_name']}, {store['home_page']}")

    #     store_index = input('Please choose the store index:')

    #     store_name = store_config['stores'][int(store_index)]['store_name']
    #     home_page = store_config['stores'][int(store_index)]['home_page']
    #     open_page_by_url(home_page)
    #     scan_store(store_name=store_name, must_include_word=setting_must_include_word, max_scroll_page=max_scroll_page)
