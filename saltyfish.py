# -*- coding: utf-8 -*-

import logging
from operator import itemgetter
import os
import pathlib
import random
import re
import shutil
import string
import sys
import time
from datetime import datetime, date, timedelta

import colorlog
import uiautomator2 as u2
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor

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

d = u2.connect("AUE66HL7XWIVJRSS")

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
    Excel单元格转换为行列索引，例如A1转换为1,1  C4转换为3,4
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
    write_path = os.getcwd()
    if not os.path.exists(write_path):
        os.makedirs(write_path)
    wb = Workbook()
    sheet = wb.active
    sheet_name = 'Sheet1'
    sheet.title = sheet_name
    sheet['A1'] = '标题'
    sheet['B1'] = '价格'
    start_row = 2
    sorted_data_list = sorted(data_list, key=itemgetter('amount'))
    mindata = min(data_list, key=itemgetter('amount'))
    output_file = os.path.join(write_path, f"{dt}-{keyword}-{mindata['amount']}.xlsx")
    for index, data in enumerate(sorted_data_list):
        sheet["A" + str(index + start_row)] = data['title']
        sheet["B" + str(index + start_row)] = data['amount']
    wb.save(filename=output_file)
    return output_file

def swipe_up():
    d.swipe_ext('up', 1)

def del_temp_file():
    if os.path.exists('images'):
        shutil.rmtree('images')

def open_page_by_keyword(keyword):
    TimeUtil.random_sleep()
    d(resourceId="com.taobao.idlefish:id/title").click()
    d.send_keys(keyword, clear=True)
    d.press('enter')

def generate_random_string(length):
    letters_and_digits = string.ascii_letters + string.digits
    return ''.join(random.choice(letters_and_digits) for i in range(length))

def get_amount(s):
    match = re.search(r'¥(\d+\.?\d*)', s)
    if match:
        amount = match.group(1)
        return float(amount)

def remove_unicode(text):
    return text.replace('\n', '')

def get_list_data(must_include_word):
    result = []
    TimeUtil.random_sleep()
    view_list = d.xpath('//android.widget.ScrollView//android.view.View').all()
    if len(view_list) > 0:
        index = 0
        for el in view_list:
            if len(el.elem.getchildren()) > 0:
                el_description = remove_unicode(str(el.attrib['content-desc']))
                for child in el.elem.getchildren():
                    el_description = f"{el_description},{remove_unicode(str(child.attrib['content-desc']))}" #combine el_description
                print(f"{index} - {el_description}")
                if must_include_word in el_description:
                    amount = get_amount(el_description)
                    if amount is not None and amount != '':
                        result.append({
                            'title': el_description,
                            'amount': amount,
                     })
            index += 1
    return result

def main_exit():
    d.set_fastinput_ime(False)
    d.app_stop(package_name)

def execute(keyword, must_include_word, max_scroll_page):
    try:
        del_temp_file()
        logger.info(d.info)
        d.app_stop(package_name)
        d.app_start(package_name, activity_name, wait=True)
        outputs = []

        logger.info(f"正在获取【{keyword}】关键字信息...")
        open_page_by_keyword(keyword)
        for i in range(max_scroll_page):
            logger.info(f"正在滑动[{i}/{max_scroll_page}]...")
            list_data = get_list_data(must_include_word)
            if list_data:
                outputs.extend(list_data)
            swipe_up()

        output_file = to_excel(outputs, keyword)
        logger.info(f"运行完成，文件路径{output_file}")
    except Exception as e:
        print(e)
        logger.error("程序运行异常:" + str(e.args[0]))
    finally:
        print("执行结束!")
        # main_exit()

if __name__ == '__main__':
    keyword = 'J老师精听精讲'
    must_include_word = 'J老师'
    max_scroll_page = 5  # 向上滑动次数
    execute(keyword=keyword, must_include_word=must_include_word, max_scroll_page=max_scroll_page)
