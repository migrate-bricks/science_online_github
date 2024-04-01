# -*- coding: utf-8 -*-
# before using this script, please refer to d = u2.connect("AUE66HL7XWIVJRSS") Change it based on 'adb devices'

import logging
from operator import itemgetter
import os
import pathlib
import random
import re
import sys
import time
from datetime import datetime, date, timedelta

import colorlog
import uiautomator2 as u2
from openpyxl import Workbook

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

console_handler = logging.StreamHandler()
console_handler.setLevel(logging.DEBUG)

formatter = colorlog.ColoredFormatter(
    fmt='%(log_color)s %(asctime)s %(levelname)s:%(message)s%(reset)s',
    datefmt="%Y-%m-%d %H:%M:%S",
    log_colors={
        'DEBUG': 'cyan',
        'INFO': 'green',
        'WARNING': 'yellow',
        'ERROR': 'red',
        'CRITICAL': 'red,bg_white',
    }
)

console_handler.setFormatter(formatter)
logger.addHandler(console_handler)

d = u2.connect("AUE66HL7XWIVJRSS") #TODO: Change it based on 'adb devices'

package_name = "com.taobao.idlefish"
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

    @staticmethod
    def tomorrow_date():
        today = date.today()
        tomorrow = today + timedelta(days=1)
        return tomorrow.strftime('%Y-%m-%d')

def get_desktop_path():
    if sys.platform == 'win32':
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    elif sys.platform == 'darwin':
        desktop_path = str(pathlib.Path.home() / "Desktop")
    else:
        desktop_path = None
    return desktop_path

def excel_cell_to_index(cell_str):
    """
    Excel cells are converted to row and column indexes, for example, A1 is converted to 1,1 and C4 is converted to 3,4
    :param cell_str:
    :return:
    """
    letter = cell_str[0]
    number = int(cell_str[1:])
    column = ord(letter.upper()) - ord('A') + 1
    row = number
    return row, column

def to_excel(data_list, keyword):
    dt = TimeUtil.curr_date()
    write_path = os.path.join(os.getcwd(), 'save')
    if not os.path.exists(write_path):
        os.makedirs(write_path)
    wb = Workbook()
    sheet = wb.active
    sheet.column_dimensions["A"].width = 100
    sheet_name = 'Sheet1'
    sheet.title = sheet_name
    sheet['A1'] = 'Title'
    sheet['B1'] = 'Price'
    sheet['C1'] = 'Wanted'
    sheet['D1'] = 'Profit'
    start_row = 2
    sorted_data_list = sorted(data_list, key=itemgetter('price'))
    mindata = min(data_list, key=itemgetter('price'))
    output_file = os.path.join(write_path, f"{dt}-{keyword}-{mindata['price']}.xlsx")
    for index, data in enumerate(sorted_data_list):
        sheet["A" + str(index + start_row)] = data['title']
        sheet["B" + str(index + start_row)] = data['price']
        sheet["C" + str(index + start_row)] = data['wanted']
        sheet["D" + str(index + start_row)] = data['price'] * data['wanted']
    wb.save(filename=output_file)
    return output_file

def swipe_up():
    d.swipe_ext('up', 0.9)

def open_page_by_keyword(keyword):
    TimeUtil.random_sleep()
    d(resourceId="com.taobao.idlefish:id/title").click()
    d.send_keys(keyword, clear=True)
    d.press('enter')

def get_price(s):
    match = re.search(r'¥(\d+\.?\d*)', s)
    if match:
        price = match.group(1)
        return float(price)

def get_wanted(s):
    match = re.search(r'(\d+\.?\d*)人想要', s)
    if match:
        price = match.group(1)
        return float(price)
    return 0

def clean_text(text):
    return text.replace('\n', '@')

def main_complete():
    d.set_fastinput_ime(False)

def execute_scan_all(keyword, must_include_word, max_scroll_page):
    try:
        logger.info(d.info)
        d.app_start(package_name, activity_name, wait=True)
        logger.info(f"Retrieving【{keyword}】products information...")
        results = []
        open_page_by_keyword(keyword)
        for i in range(max_scroll_page):
            logger.info(f"Scrolling to [{i}/{max_scroll_page}] page...")
            TimeUtil.random_sleep()
            view_list = d.xpath('//android.widget.ScrollView//android.view.View').all()
            if len(view_list) > 0:
                for el in view_list:
                    if len(el.elem.getchildren()) > 0:
                        el_description = clean_text(str(el.attrib['content-desc']))
                        for child in el.elem.getchildren():
                            el_description = f"{el_description}|{clean_text(str(child.attrib['content-desc']))}" #combine el_description
                        if must_include_word in el_description:
                            price = get_price(el_description)
                            wanted = get_wanted(el_description)
                            if price is not None and price != '' and not any(d['title'] == el_description for d in results): # skip duplicated item
                                logger.info(f"【{len(results)+1}】-description:{el_description}, price:{price}, wanted:{wanted}")
                                results.append({ 'title': el_description, 'price': price, 'wanted': wanted})
            if d(descriptionContains='到底了').exists: # alread on the end of the page
                break
            swipe_up()
        output_file = to_excel(results, keyword)
        logger.info(f"Execution completed, file path: {output_file}")
    except Exception as e:
        print(e)
        logger.error("Program runs Error:" + str(e.args[0]))
    finally:
        main_complete()
        print("Execution Completed!")

if __name__ == '__main__':
    keyword = 'tanner老师'
    must_include_word = 'J老师'
    max_scroll_page = 100
    execute_scan_all(keyword=keyword, must_include_word=must_include_word, max_scroll_page=max_scroll_page)
